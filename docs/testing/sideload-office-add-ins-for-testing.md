---
title: Charger une version test des compléments Office dans Office sur le Web
description: Testez votre complément Office dans Office sur le Web en chargeant la version test.
ms.date: 09/02/2022
ms.localizationpriority: medium
ms.openlocfilehash: 128e3537ac0ece5b7574dfec6d9d5c67b8d95a7b
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810381"
---
# <a name="sideload-office-add-ins-to-office-on-the-web"></a>Charger une version test des compléments Office dans Office sur le Web

Lorsque vous chargez une version test d’un complément, vous pouvez installer le complément sans le placer au préalable dans un catalogue de compléments. Cela est utile lors du test et du développement de votre complément, car vous pouvez voir comment votre complément s’affichera et fonctionnera.

Lorsque vous chargez une version test d’un complément sur le web, le manifeste du complément est stocké dans le stockage local du navigateur. Par conséquent, si vous effacez le cache du navigateur ou basculez vers un autre navigateur, vous devez charger à nouveau la version test du complément.

Les étapes de chargement indépendant d’un complément sur le web varient en fonction des facteurs suivants.

- L’application hôte (par exemple, Excel, Word, Outlook)
- Quel outil a créé le projet de complément (par exemple, Visual Studio, générateur Yeoman pour les compléments Office, ou aucun des deux)
- Si vous effectuez un chargement indépendant vers Office sur le Web avec un compte Microsoft ou un compte dans un locataire Microsoft 365

Dans la liste suivante, accédez à la section ou à l’article qui correspond à votre scénario. Notez que le premier scénario de la liste s’applique aux compléments Outlook. Les autres scénarios s’appliquent aux compléments non-Outlook.

- Si vous rechargez une version test d’un complément Outlook, consultez l’article Chargement indépendant des [compléments Outlook à des fins de test](../outlook/sideload-outlook-add-ins-for-testing.md).
- Si vous avez créé le complément à l’aide du [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md), consultez [Charger une version test d’un complément créé par Yeoman dans Office sur le Web](#sideload-a-yeoman-created-add-in-to-office-on-the-web).
- Si vous avez créé le complément à l’aide de Visual Studio, consultez [Charger une version test d’un complément sur le web lors de l’utilisation de Visual Studio](#sideload-an-add-in-on-the-web-when-using-visual-studio).
- Pour tous les autres cas, consultez l’une des sections suivantes.

  - Si vous effectuez un chargement indépendant vers Office sur le Web avec un compte Microsoft, consultez [Charger manuellement une version test d’un complément pour Office sur le Web](#manually-sideload-an-add-in-to-office-on-the-web).
  - Si vous effectuez un chargement indépendant pour Office sur le Web avec un compte dans un locataire Microsoft 365, consultez [Charger une version test d’un complément à Microsoft 365](#sideload-an-add-in-to-microsoft-365).

## <a name="sideload-a-yeoman-created-add-in-to-office-on-the-web"></a>Charger une version test d’un complément créé par Yeoman dans Office sur le Web

Ce processus est pris en charge pour **Excel**, **OneNote**, **PowerPoint** et **Word** uniquement. Cet exemple de projet suppose que vous utilisez un projet créé avec le [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md).

1. Ouvrez [Office sur le Web](https://office.live.com/) ou OneDrive. À l’aide de l’option **Créer** , créez un document dans **Excel**, **OneNote**, **PowerPoint** ou **Word**. Dans ce nouveau document, sélectionnez **Partager**, **Copier le lien**, puis copiez l’URL.

1. Dans la ligne de commande commençant dans le répertoire racine de votre projet, exécutez la commande suivante. Remplacez « {url} » par l’URL que vous avez copiée.

    [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

1. La première fois que vous utilisez cette méthode pour charger une version test d’un complément sur le web, vous voyez une boîte de dialogue vous demandant d’activer le mode développeur. Cochez la case **Activer le mode développeur maintenant** , puis sélectionnez **OK**.

1. Une deuxième boîte de dialogue s’affiche, vous demandant si vous souhaitez inscrire un manifeste de complément Office à partir de votre ordinateur. Sélectionnez **Oui**.

1. Votre complément est installé. S’il a une commande de complément, elle doit apparaître sur le ruban ou le menu contextuel. S’il s’agit d’un complément du volet Office sans commande de complément, le volet Office doit apparaître.

## <a name="sideload-an-add-in-on-the-web-when-using-visual-studio"></a>Charger une version test d’un complément sur le web lors de l’utilisation de Visual Studio

Si vous utilisez Visual Studio pour développer votre complément, appuyez sur **F5** pour ouvrir un document Office dans office de *bureau* , créer un document vide et charger une version test du complément. Lorsque vous souhaitez charger une version test vers *Office sur le Web*, le processus de chargement indépendant est similaire au chargement indépendant manuel sur le web. La seule différence est que vous devez mettre à jour la valeur de l’élément **SourceURL** , et éventuellement d’autres éléments, dans votre manifeste pour inclure l’URL complète où le complément est déployé.

1. Dans Visual Studio, choisissez **Fenêtre Afficher les** > **propriétés**.

1. Dans l’**Explorateur de solutions**, sélectionnez le projet web. Cela affiche les propriétés du projet dans la fenêtre **Propriétés** .

1. Dans la fenêtre Propriétés, copiez l’**URL SSL**.

1. Dans le projet de complément, ouvrez le fichier XML de manifeste. Veillez à modifier le code XML source. Pour certains types de projets, Visual Studio ouvre une vue visuelle du code XML qui ne fonctionne pas pour l’étape suivante.

1. Cherchez toutes les instances de **~remoteAppUrl/** et remplacez-les par l’URL SSL que vous venez de copier. Vous verrez plusieurs remplacements en fonction du type de projet, et les nouvelles URL ressembleront à `https://localhost:44300/Home.html`.

1. **Enregistrez** le fichier XML.

1. Dans le **Explorateur de solutions**, ouvrez le menu contextuel du projet web (par exemple, en cliquant dessus avec le bouton droit), puis choisissez **Déboguer** > **Démarrer une nouvelle instance**. Cela exécute le projet web sans lancer Office.

1. À partir Office sur le Web, chargez une version test du complément en suivant les étapes décrites dans [Charger manuellement une version test d’un complément dans Office sur le Web](#manually-sideload-an-add-in-to-office-on-the-web).

## <a name="manually-sideload-an-add-in-to-office-on-the-web"></a>Charger manuellement une version test d’un complément dans Office sur le Web

Cette méthode n’utilise pas la ligne de commande et peut être effectuée à l’aide de commandes uniquement dans l’application hôte (par exemple, Excel).

1. Ouvrez [Office sur le Web](https://office.com/). Ouvrez un document dans **Excel**, **OneNote**, **PowerPoint** ou  **Word**. 

1. Sous l’onglet **Insertion** , dans la section **Compléments** , choisissez **Compléments Office**.

1. Dans la boîte de dialogue **Compléments Office** , sélectionnez l’onglet **MES COMPLÉMENTS** , choisissez **Gérer mes compléments**, puis **Charger mon complément**.

    ![La boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit « Gérer mes compléments » et une liste déroulante en dessous avec l’option « Charger mon complément ».](../images/office-add-ins-my-account.png)

1. **Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.

    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

1. Vérifiez que votre complément est installé. Par exemple, s’il a une commande de complément, elle doit apparaître sur le ruban ou le menu contextuel. S’il s’agit d’un complément du volet Office qui n’a pas de commandes de complément, le volet Office doit apparaître.

> [!NOTE]
> Pour tester votre complément Office avec Microsoft Edge avec le WebView d’origine (EdgeHTML), une étape de configuration supplémentaire est nécessaire. Dans une invite de commandes Windows, exécutez la ligne suivante : `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes`. Cela n’est pas obligatoire lorsqu’Office utilise edge WebView2 basé sur Chromium. Pour plus d’informations, voir [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).

[!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

## <a name="sideload-an-add-in-to-microsoft-365"></a>Charger une version test d’un complément dans Microsoft 365

1. Connectez-vous à votre compte Microsoft 365.

1. Ouvrez le lanceur d’applications à l’extrémité gauche de la barre d’outils, sélectionnez **Excel**, **OneNote**, **PowerPoint** ou **Word**, puis créez un document.

1. Sous l’onglet **Insertion** , sélectionnez le bouton **Compléments** .

1. Suivez les étapes 3 à 5 de la section [Charger manuellement une version test d’un complément dans Office sur le Web](#manually-sideload-an-add-in-to-office-on-the-web).

## <a name="remove-a-sideloaded-add-in"></a>Supprimer un complément chargé de manière indépendante

Pour supprimer un complément chargé de manière indépendante dans Office sur le Web, effacez simplement le cache de votre navigateur. Si vous apportez des modifications au manifeste de votre complément (par exemple, mettre à jour les noms de fichiers des icônes ou le texte des commandes de complément), vous devrez peut-être effacer le cache de votre navigateur, puis recharger le complément à l’aide du manifeste mis à jour. Cela permet Office sur le Web d’afficher le complément tel qu’il est décrit par le manifeste mis à jour.

## <a name="see-also"></a>Voir aussi

- [Chargement de versions test de compléments Office sur Mac](sideload-an-office-add-in-on-mac.md)
- [Chargement de versions test de compléments Office sur iPad](sideload-an-office-add-in-on-ipad.md)
- [Chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md)
- [Vider le cache Office](clear-cache.md)
