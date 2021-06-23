---
title: Chargement de version test des compléments Office dans Office sur le web
description: Testez votre Office dans Office sur le Web chargement de version test.
ms.date: 04/14/2021
localization_priority: Normal
ms.openlocfilehash: e830ccbb6a4e325d6d70c3612492009b5e3d1570
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077217"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a>Chargement de version test des compléments Office dans Office sur le web

Lorsque vous chargez une version de version sideload d’un add-in, vous pouvez l’installer sans le placer au premier abord dans le catalogue de modules. Cela est utile lors du test et du développement de votre add-in, car vous pouvez voir comment il s’affiche et fonctionne.

Lorsque vous chargez une version de chargement d’un module, le manifeste du module est stocké dans le stockage local du navigateur. Par exemple, si vous effacer le cache du navigateur ou que vous basculez vers un autre navigateur, vous devez recharger le module.

Le chargement de version secondaire varie d’une application hôte à l’autre (par exemple, Excel).

> [!NOTE]
> Le chargement de version de version secondaire, comme décrit dans cet article, est pris en charge sur Excel, OneNote, PowerPoint et Word. Pour charger une version test de complément Outlook, voir la rubrique relative au [chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md).

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a>Chargement de version test d’un complément Office dans Office sur le web

Ce processus est pris en charge **pour Excel,** **OneNote,** **PowerPoint** et **Word** uniquement. Pour les autres applications hôtes, consultez les instructions de chargement de version de version manuelle dans la section suivante. Cet exemple de projet suppose que vous utilisez un projet créé avec le générateur [Yeoman](https://github.com/OfficeDev/generator-office)pour Office de recherche.

1. Ouvrez [Office sur le Web](https://office.live.com/). À **l’aide de l’option** Créer, créez un document **dans Excel,** **OneNote,** **PowerPoint** ou **Word.** Dans ce nouveau document, **sélectionnez Partager** dans le ruban, **sélectionnez Copier** le lien et copiez l’URL.

2. Dans le répertoire racine de vos fichiers de projet Yo Office, ouvrez **package.jsfichier on.** Dans la section **de config** de ce fichier, créez une `"document"` propriété. Collez l’URL que vous avez copiée comme valeur pour la `"document"` propriété. Par exemple, le vôtre ressemblera à ceci :

    ```json
      "config": {
        "document": "<YOUR URL>",
        ...
      }
    ```

    > [!TIP]
    > Si vous créez un add-in qui n’utilise pas notre générateur Yeoman, vous pouvez ajouter des paramètres de requête à l’URL de votre document, en ajoutez ce qui suit à l’URL existante :

    - Port du serveur dev, tel que `&wdaddindevserverport=3000` .
    - Nom du fichier manifeste, tel que `&wdaddinmanifestfile=manifest1.xml` .
    - GUID du manifeste, tel `&wdaddinmanifestguid=05c2e1c9-3e1d-406e-9a91-e9ac64854143` que .

    > Si vous utilisez le générateur Yeoman, l’ajout de ces informations n’est pas nécessaire, car les outils Yeoman ajoutent ces informations automatiquement.
    > Notez que dans les deux cas, toutefois, vous ne pouvez charger les manifestes qu’à partir de l’host local.

3. Dans la ligne de commande commençant dans le répertoire racine de votre projet, exécutez la commande suivante `npm run start:web` :

4. La première fois que vous utilisez cette méthode pour le chargement indépendant d’un application sur le web, une boîte de dialogue vous demande d’activer le mode développeur. Activez la case à cocher **activer le mode développeur maintenant** et sélectionnez **OK.**

5. Vous verrez une deuxième boîte de dialogue vous demandant si vous souhaitez inscrire un manifeste de Office à partir de votre ordinateur. Vous devez sélectionner **Oui**.

6. Votre add-in est installé. S’il s’agit d’une commande de add-in, elle doit apparaître dans le ruban ou le menu contexto. S’il s’agit d’un add-in du volet Des tâches, celui-ci doit apparaître.

## <a name="sideload-an-office-add-in-in-office-on-the-web-manually"></a>Chargement de version Office de votre Office sur le Web manuellement

Cette méthode n’utilise pas la ligne de commande et peut être accomplie à l’aide de commandes uniquement dans l’application hôte (par exemple, Excel).

1. Ouvrez [Office sur le Web](https://office.live.com/). Ouvrez un document **dans Excel,** **Word** **ou PowerPoint**. Sous **l’onglet** Insérer dans le ruban de la **section** Des Office, sélectionnez **Ajouter.**

1. Dans la **boîte Office** de dialogue Des Télécharger, sélectionnez l’onglet MES **ADD-INS,** choisissez Gérer mes **applications,** puis Télécharger **My Add-in**.

    ![La boîte de dialogue Office des applications avec une zone de texte dans le coin supérieur droit de la lecture « Gérer mes applications » et une zone de texte en dessous avec l’option « Télécharger Mon add-in ».](../images/office-add-ins-my-account.png)

1. **Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.

    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

1. Vérifiez que votre complément est installé. S’il s’agit d’une commande de complément, elle doit apparaître dans le ruban ou dans le menu contextuel. S’il s’agit d’un complément du volet Office, le volet doit apparaître.

> [!NOTE]
> Pour tester votre complément Office avec Microsoft Edge webview d’origine (EdgeHTML), une étape de configuration supplémentaire est requise. Dans une invite Windows commande, exécutez la ligne suivante `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes` : Cela n’est pas nécessaire lorsque Office utilise Chromium WebView2 Edge basé sur Chromium web. Pour plus d’informations, voir [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

## <a name="sideload-an-office-add-in"></a>Chargement de version Office un module

1. Connectez-vous à Microsoft 365 compte.

2. Ouvrez l’Lanceur application à l’extrémité gauche de la barre d’outils, sélectionnez **Excel,** **Word** ou **PowerPoint,** puis créez un document.

3. Les étapes 3 à 6 sont identiques à celles de la section précédente, **Chargement d’une version de test d’un complément Office dans Office sur le web**.

## <a name="sideload-an-add-in-when-using-visual-studio"></a>Chargement d’une version test d’un complément lors de l’utilisation de Visual Studio

Si vous utilisez Visual Studio pour développer votre application, le processus de chargement de version de version de chargement de version est similaire au chargement de version manuelle sur le web. La seule différence est que vous devez mettre à jour la valeur de l’élément **SourceURL** dans votre manifeste afin d’inclure l’URL complète de déploiement du complément.

> [!NOTE]
> Si vous pouvez charger une version test des compléments à partir de Visual Studio vers Office sur le web, vous ne pouvez pas les déboguer à partir de Visual Studio. Pour déboguer, vous devrez utiliser les outils de débogage du navigateur. Pour plus d’informations, voir [Débogage de compléments dans Office sur le web](debug-add-ins-in-office-online.md).

1. Dans Visual Studio, affichez la fenêtre **Propriétés** en choisissant **Affichage** > **Fenêtre Propriétés**.
2. Dans l’**Explorateur de solutions**, sélectionnez le projet web. Cela a pour effet d’afficher les propriétés du projet dans la fenêtre **Propriétés**.
3. Dans la fenêtre Propriétés, copiez l’**URL SSL**.
4. Dans le projet de complément, ouvrez le fichier XML de manifeste. Veillez à modifier le code XML source. Pour certains types de projets, Visual Studio ouvre un affichage visuel du code XML qui ne fonctionnera pas pour l’étape suivante.
5. Cherchez toutes les instances de **~remoteAppUrl/** et remplacez-les par l’URL SSL que vous venez de copier. Vous verrez plusieurs remplacements en fonction du type de projet, et les nouvelles URL ressembleront à `https://localhost:44300/Home.html`.
6. Enregistrez le fichier XML.
7. Cliquez avec le bouton droit sur le projet web, puis sélectionnez **Déboguer** > **Démarrer une nouvelle instance**. Cela a pour effet d’exécuter le projet web sans lancer Office.
8. À partir d’Office sur le web, chargez la version test du complément en suivant les étapes décrites précédemment dans [Chargement de version test d’un complément Office dans Office sur le web](#sideload-an-office-add-in-in-office-on-the-web).

## <a name="remove-a-sideloaded-add-in"></a>Supprimer un add-in chargé de nouveau

Vous pouvez supprimer un add-in précédemment chargé de nouveau en effantant le cache de votre navigateur. Si vous modifiez le manifeste de votre add-in (par exemple, mettez à jour les noms de fichiers des icônes ou le texte des commandes de votre module), vous devrez peut-être effacer le [cache Office,](clear-cache.md) puis recharger le module à l’aide du manifeste mis à jour. Cette action permettra à Office d’afficher le complément tel que décrit par le manifeste mis à jour.

## <a name="see-also"></a>Voir aussi

- [Chargement de version test de compléments Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
- [Chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md)
- [Vider le cache Office](clear-cache.md)
