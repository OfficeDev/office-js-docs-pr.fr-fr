---
title: Publier un complément à l’aide de Visual Studio Code et d’Azure
description: Comment publier un complément à l’aide de Visual Studio Code et d’Azure Active Directory
ms.date: 09/07/2022
ms.custom: vscode-azure-extension-update-completed
ms.localizationpriority: medium
ms.openlocfilehash: f1e931042f7796333bbc75b2666f3564e63b9f7b
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467193"
---
# <a name="publish-an-add-in-developed-with-visual-studio-code"></a>Publier un complément développé avec Visual Studio Code

Cet article explique comment un complément Office que vous avez créé à l’aide du générateur Yeoman et développé avec [Visual Studio Code (VS Code)](https://code.visualstudio.com) ou un autre éditeur.

> [!NOTE]
> Pour plus d’informations sur la publication d’un complément Office que vous avez créé à l’aide de Visual Studio, voir [Publier votre complément à l’aide de Visual Studio](package-your-add-in-using-visual-studio.md).

## <a name="publishing-an-add-in-for-other-users-to-access"></a>Publication d’un complément pour accéder à d’autres utilisateurs

An Office Add-in consists of a web application and a manifest file. The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.

Pendant le développement, vous pouvez exécuter le complément sur votre serveur web local (`localhost`). Lorsque vous êtes prêt à la publier pour que d’autres utilisateurs y accèdent, vous devez déployer l’application web et mettre à jour le manifeste pour spécifier l’URL de l’application déployée.

Lorsque votre complément fonctionne comme vous le souhaitez, vous pouvez le publier directement via Visual Studio Code à l’aide de l’extension stockage Azure.

## <a name="using-visual-studio-code-to-publish"></a>Utilisation de Visual Studio Code pour publier

>[!NOTE]
> Ces étapes fonctionnent uniquement pour les projets créés avec le générateur Yeoman.

1. Ouvrez votre projet à partir de son dossier racine dans Visual Studio Code (VS Code).
1. Sélectionnez **Afficher > extensions** (Ctrl+Maj+X) pour ouvrir la vue Extensions.
1. Recherchez l’extension **Stockage Azure** et installez-la.
1. Une fois installée, une icône Azure est ajoutée à la **barre d’activités**. Sélectionnez-la pour accéder à l’extension. Si la **barre d’activité** est masquée, ouvrez-la en sélectionnant **Afficher > Apparence > Barre d’activité**.
1. Sélectionnez **Se connecter à Azure** pour vous connecter à votre compte Azure. Si vous n’avez pas encore de compte Azure, créez-en un en sélectionnant **Créer un compte Azure**. Suivez les étapes fournies pour configurer votre compte.

    :::image type="content" source="../images/azure-extension-sign-in.png" alt-text="Bouton Se connecter à Azure sélectionné dans l’extension Azure.":::

1. Une fois connecté, vos comptes de stockage Azure s’affichent dans l’extension. Si vous n’avez pas encore de compte de stockage, créez-en un à l’aide de l’option **Créer un compte de stockage** dans la palette de commandes. Nommez votre compte de stockage sous un nom global unique, en utilisant uniquement « a-z » et « 0-9 ». Notez que par défaut, cela crée un compte de stockage et un groupe de ressources portant le même nom. Il place automatiquement le compte de stockage dans la région USA Ouest. Cela peut être ajusté en ligne via [votre compte Azure](https://portal.azure.com/).

    :::image type="content" source="../images/azure-extension-create-storage-account.png" alt-text="Sélection de comptes de stockage > créer un compte de stockage dans l’extension Azure.":::

1. Cliquez avec le bouton droit sur votre compte de stockage et **sélectionnez Configurer le site web statique**. Vous serez invité à entrer le nom du document d’index et le nom du document 404. Remplacez le nom du document d’index par défaut `index.html` **`taskpane.html`** par . Vous pouvez également modifier le nom du document 404, mais ce n’est pas obligatoire.
1. Cliquez à nouveau avec le bouton droit sur votre compte de stockage, puis **sélectionnez Parcourir le site web statique**. Dans la fenêtre du navigateur qui s’ouvre, copiez l’URL du site web.
1. Ouvrez le fichier manifeste de votre projet (`manifest.xml`) et remplacez toutes les références à votre URL localhost (par exemple `https://localhost:3000`) par l’URL que vous avez copiée. Ce point de terminaison est l’URL du site web statique pour votre compte de stockage nouvellement créé. Enregistrez les modifications apportées à votre fichier manifeste.
1. Ouvrez une invite de ligne de commande ou une fenêtre de terminal et accédez au répertoire racine de votre projet de complément. Exécutez la commande suivante pour préparer tous les fichiers pour le déploiement de production.

    ```command&nbsp;line
    npm run build
    ```

    Une fois la build terminée, le dossier **dist** dans le répertoire racine de votre projet de complément contient les fichiers que vous allez déployer dans les étapes suivantes.

1. Dans VS Code, accédez à l’Explorateur et cliquez avec le bouton droit sur le dossier **dist** , puis sélectionnez **Déployer sur le site web statique via stockage Azure**. Lorsque vous y êtes invité, sélectionnez le compte de stockage que vous avez créé précédemment.

    :::image type="content" source="../images/deploy-to-static-website.png" alt-text="Sélectionnez le dossier dist, cliquez avec le bouton droit, puis sélectionnez Déployer sur le site web statique via Stockage Azure.":::

1. Une fois le déploiement terminé, cliquez avec le bouton droit sur le compte de stockage que vous avez créé précédemment, puis **sélectionnez Parcourir le site web statique**. Le site web statique s’ouvre et affiche le volet Office.

1. Enfin, [chargez de manière indépendante le fichier manifeste](../testing/sideload-office-add-ins-for-testing.md) et le complément sera chargé à partir du site web statique que vous venez de déployer.

## <a name="deploy-custom-functions-for-excel"></a>Déployer des fonctions personnalisées pour Excel

Si votre complément a des fonctions personnalisées, il existe quelques étapes supplémentaires pour les activer sur le compte de stockage Azure. Tout d’abord, activez CORS pour qu’Office puisse accéder au fichier functions.json.

1. Cliquez avec le bouton droit sur le compte de stockage Azure, puis sélectionnez **Ouvrir dans le portail**.
1. Dans le groupe Paramètres, sélectionnez **Partage des ressources (CORS).** Vous pouvez également utiliser la zone de recherche pour trouver cette option.
1. Créez une règle CORS avec les paramètres suivants.

    |Propriété        |Valeur                        |
    |----------------|-----------------------------|
    |Origines autorisées | \*                          |
    |Méthodes autorisées | GET                         |
    |En-têtes autorisés | \*                          |
    |En-têtes exposés | Access-Control-Allow-Origin |
    |Âge maximal         | 200                          |

1. Sélectionnez **Enregistrer**.

> [!CAUTION]
> Cette configuration CORS suppose que tous les fichiers de votre serveur sont disponibles publiquement pour tous les domaines.  

Ensuite, ajoutez un type MIME pour les fichiers JSON.

1. Créez un fichier dans le dossier /src nommé **web.config**.
1. Insérez le code XML suivant et enregistrez le fichier.

    ```xml
    <?xml version="1.0"?>
    <configuration>
      <system.webServer>
        <staticContent>
          <mimeMap fileExtension=".json" mimeType="application/json" />
        </staticContent>
      </system.webServer>
    </configuration> 
    ```

1. Ouvrez le fichier **webpack.config.js**.
1. Ajoutez le code suivant dans la liste des éléments permettant de `plugins` copier le web.config dans le bundle lors de l’exécution de la build.

    ```javascript
    new CopyWebpackPlugin({
      patterns: [
      {
        from: "src/web.config",
        to: "src/web.config",
      },
     ],
    }),
    ```

1. Ouvrez une invite de ligne de commande et accédez au répertoire racine de votre projet de complément. Exécutez ensuite la commande suivante pour préparer tous les fichiers pour le déploiement.

    ```command&nbsp;line
    npm run build
    ```

    Une fois la génération terminée, le dossier **dist** dans le répertoire racine de votre projet de complément contient les fichiers que vous allez déployer.

1. Pour déployer, dans **l’Explorateur** VS Code, cliquez avec le bouton droit sur le dossier **dist** , puis sélectionnez **Déployer sur le site web statique via stockage Azure**. Lorsque vous y êtes invité, sélectionnez le compte de stockage que vous avez créé précédemment. Si vous avez déjà déployé le dossier **dist** , vous êtes invité à remplacer les fichiers du stockage Azure par les dernières modifications.

## <a name="see-also"></a>Voir aussi

- [Développement d’un complément Office avec Visual Studio Code](../develop/develop-add-ins-vscode.md)
- [Déploiement et publication de votre complément Office](../publish/publish.md)
- [Prise en charge du partage de ressources cross-origin (CORS) pour Stockage Azure](/rest/api/storageservices/cross-origin-resource-sharing--cors--support-for-the-azure-storage-services)
