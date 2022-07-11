---
title: Chargement de version test des compléments Office dans Office sur le web
description: Testez votre complément Office dans Office sur le Web par chargement indépendant.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 32d80a10ccddab93fc8d41151be6a2842d3732cb
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713027"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a>Chargement de version test des compléments Office dans Office sur le web

Lorsque vous chargez de manière indépendante un complément, vous pouvez installer le complément sans le placer d’abord dans le catalogue de compléments. Cela est utile lors du test et du développement de votre complément, car vous pouvez voir comment votre complément apparaîtra et fonctionnera.

Lorsque vous chargez de manière indépendante un complément, le manifeste du complément est stocké dans le stockage local du navigateur. Par conséquent, si vous effacez le cache du navigateur ou basculez vers un autre navigateur, vous devez recharger le complément.

Le chargement indépendant varie d’une application hôte à l’autre (par exemple, Excel).

> [!NOTE]
> Le chargement indépendant, comme décrit dans cet article, est pris en charge sur Excel, OneNote, PowerPoint et Word. Pour charger une version test de complément Outlook, voir la rubrique relative au [chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md).

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a>Chargement de version test d’un complément Office dans Office sur le web

Ce processus est pris en charge uniquement pour **Excel**, **OneNote**, **PowerPoint** et **Word** . Pour les autres applications hôtes, consultez les instructions de chargement indépendant manuelles dans la section suivante. Cet exemple de projet suppose que vous utilisez un projet créé avec le [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md).

1. Ouvrez [Office sur le Web](https://office.live.com/). À l’aide de l’option **Créer** , créez un document dans **Excel**, **OneNote**, **PowerPoint** ou **Word**. Dans ce nouveau document, **sélectionnez Partager** dans le ruban, sélectionnez **Copier le lien** et copiez l’URL.

1. Dans le répertoire racine de vos fichiers projet yo Office, ouvrez le fichier **package.json** . Dans la section **config** de ce fichier, créez une `"document"` propriété. Collez l’URL que vous avez copiée comme valeur de la `"document"` propriété. Par exemple, la vôtre ressemble à ceci :

    ```json
      "config": {
        "document": "<YOUR URL>",
        ...
      }
    ```

    > [!TIP]
    > Si vous créez un complément qui n’utilise pas notre générateur Yeoman, vous pouvez ajouter des paramètres de requête à l’URL de votre document, en ajoutant ce qui suit à l’URL existante.
    >
    > - Port du serveur de développement, par `&wdaddindevserverport=3000`exemple .
    > - Nom du fichier manifeste, par `&wdaddinmanifestfile=manifest1.xml`exemple .
    > - GUID du manifeste, tel que `&wdaddinmanifestguid=05c2e1c9-3e1d-406e-9a91-e9ac64854143`.
    >
    > Si vous utilisez le générateur Yeoman, l’ajout de ces informations n’est pas nécessaire, car les outils Yeoman ajoutent ces informations automatiquement.
    > Notez que dans les deux cas, toutefois, vous pouvez uniquement charger des manifestes à partir de localhost.

1. Dans la ligne de commande commençant au répertoire racine de votre projet, exécutez la commande suivante. Remplacez « {url} » par l’URL d’un document Office sur votre OneDrive ou une bibliothèque SharePoint à laquelle vous disposez d’autorisations.

    [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

1. La première fois que vous utilisez cette méthode pour charger une version test d’un complément sur le web, vous verrez une boîte de dialogue vous demandant d’activer le mode développeur. Activez la case à cocher **Activer le mode développeur,** puis sélectionnez **OK**.

1. Vous verrez une deuxième boîte de dialogue vous demandant si vous souhaitez inscrire un manifeste de complément Office à partir de votre ordinateur. Vous devez sélectionner **Oui**.

1. Votre complément est installé. S’il s’agit d’une commande de complément, elle doit apparaître sur le ruban ou le menu contextuel. S’il s’agit d’un complément du volet Office, le volet Office doit apparaître.

## <a name="sideload-an-office-add-in-in-office-on-the-web-manually"></a>Charger manuellement un complément Office dans Office sur le Web

Cette méthode n’utilise pas la ligne de commande et peut être effectuée à l’aide de commandes uniquement au sein de l’application hôte (par exemple, Excel).

1. Ouvrez [Office sur le Web](https://office.com/). Ouvrez un document dans **Excel**, **OneNote**, **PowerPoint** ou  **Word**. Sous l’onglet **Insertion** du ruban dans la section **Compléments** , choisissez **Compléments Office**.

1. Dans la boîte **de dialogue Compléments Office** , sélectionnez l’onglet **MES COMPLÉMENTS** , choisissez **Gérer mes compléments**, puis **chargez mon complément**.

    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit en lisant « Gérer mes compléments » et une liste déroulante en dessous avec l’option « Télécharger mon complément ».](../images/office-add-ins-my-account.png)

1. **Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.

    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

1. Vérifiez que votre complément est installé. S’il s’agit d’une commande de complément, elle doit apparaître dans le ruban ou dans le menu contextuel. S’il s’agit d’un complément du volet Office, le volet doit apparaître.

> [!NOTE]
> Pour tester votre complément Office avec Microsoft Edge avec le WebView d’origine (EdgeHTML), une étape de configuration supplémentaire est requise. Dans une invite de commandes Windows, exécutez la ligne suivante : `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes`. Cela n’est pas obligatoire lorsqu’Office utilise edge WebView2 basé sur Chromium. Pour plus d’informations, voir [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).

[!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

## <a name="sideload-an-office-add-in"></a>Charger une version test d’un complément Office

1. Connectez-vous à votre compte Microsoft 365.

1. Ouvrez le lanceur d’applications à l’extrémité gauche de la barre d’outils, sélectionnez **Excel**, **PowerPoint** ou **Word**, puis créez un document.

1. Les étapes 3 à 6 sont identiques à celles de la section précédente, **Chargement d’une version de test d’un complément Office dans Office sur le web**.

## <a name="sideload-an-add-in-when-using-visual-studio"></a>Chargement d’une version test d’un complément lors de l’utilisation de Visual Studio

Si vous utilisez Visual Studio pour développer votre complément, le processus de chargement indépendant est similaire au chargement indépendant manuel sur le web. La seule différence est que vous devez mettre à jour la valeur de l’élément **SourceURL** dans votre manifeste afin d’inclure l’URL complète de déploiement du complément.

> [!NOTE]
> Si vous pouvez charger une version test des compléments à partir de Visual Studio vers Office sur le web, vous ne pouvez pas les déboguer à partir de Visual Studio. Pour déboguer, vous devrez utiliser les outils de débogage du navigateur. Pour plus d’informations, voir [Débogage de compléments dans Office sur le web](debug-add-ins-in-office-online.md).

1. Dans Visual Studio, affichez la fenêtre **Propriétés** en choisissant **Affichage** > **Fenêtre Propriétés**.
1. Dans l’**Explorateur de solutions**, sélectionnez le projet web. Cela a pour effet d’afficher les propriétés du projet dans la fenêtre **Propriétés**.
1. Dans la fenêtre Propriétés, copiez l’**URL SSL**.
1. Dans le projet de complément, ouvrez le fichier XML de manifeste. Veillez à modifier le code XML source. Pour certains types de projets, Visual Studio ouvre un affichage visuel du code XML qui ne fonctionnera pas pour l’étape suivante.
1. Cherchez toutes les instances de **~remoteAppUrl/** et remplacez-les par l’URL SSL que vous venez de copier. Vous verrez plusieurs remplacements en fonction du type de projet, et les nouvelles URL ressembleront à `https://localhost:44300/Home.html`.
1. Enregistrez le fichier XML.
1. Cliquez avec le bouton droit sur le projet web, puis sélectionnez **Déboguer** > **Démarrer une nouvelle instance**. Cela a pour effet d’exécuter le projet web sans lancer Office.
1. À partir d’Office sur le web, chargez la version test du complément en suivant les étapes décrites précédemment dans [Chargement de version test d’un complément Office dans Office sur le web](#sideload-an-office-add-in-in-office-on-the-web).

## <a name="remove-a-sideloaded-add-in"></a>Supprimer un complément chargé de manière indépendante

Vous pouvez supprimer un complément précédemment chargé de manière indépendante en désactivant le cache de votre navigateur. Si vous apportez des modifications au manifeste de votre complément (par exemple, mettez à jour les noms de fichiers d’icônes ou le texte des commandes de complément), vous devrez peut-être effacer le cache de votre navigateur, puis recharger le complément à l’aide du manifeste mis à jour. Cela permet à Office sur le Web d’afficher le complément tel qu’il est décrit par le manifeste mis à jour.

## <a name="see-also"></a>Voir aussi

- [Chargement indépendant des compléments Office sur Mac](sideload-an-office-add-in-on-mac.md)
- [Chargement indépendant des compléments Office sur iPad](sideload-an-office-add-in-on-ipad.md)
- [Chargement de version test des compléments Outlook pour les tester](../outlook/sideload-outlook-add-ins-for-testing.md)
- [Vider le cache Office](clear-cache.md)
