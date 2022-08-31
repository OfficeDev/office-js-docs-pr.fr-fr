---
title: Publier un complément à l’aide de Visual Studio Code et d’Azure
description: Comment publier un complément à l’aide de Visual Studio Code et d’Azure Active Directory
ms.date: 08/19/2022
ms.custom: vscode-azure-extension-update-completed
ms.localizationpriority: medium
ms.openlocfilehash: 1c82d62e9f92453839084179d7ef9e0a8e2c8ca3
ms.sourcegitcommit: eef2064d7966db91f8401372dd255a32d76168c2
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/31/2022
ms.locfileid: "67464784"
---
# <a name="publish-an-add-in-developed-with-visual-studio-code"></a>Publier un complément développé avec Visual Studio Code

Cet article explique comment un complément Office que vous avez créé à l’aide du générateur Yeoman et développé avec [Visual Studio Code (VS Code)](https://code.visualstudio.com) ou un autre éditeur.

> [!NOTE]
> Pour plus d’informations sur la publication d’un complément Office que vous avez créé à l’aide de Visual Studio, voir [Publier votre complément à l’aide de Visual Studio](package-your-add-in-using-visual-studio.md).

## <a name="publishing-an-add-in-for-other-users-to-access"></a>Publication d’un complément pour accéder à d’autres utilisateurs

Un complément Office se compose d’une application et d’un fichier manifeste. L’application Web définit l’interface utilisateur et les fonctionnalités du complément, tandis que le manifeste spécifie l’emplacement de l’application Web et définit les paramètres et fonctionnalités du complément.

Pendant le développement, vous pouvez exécuter le complément sur votre serveur web local (`localhost`). Lorsque vous êtes prêt à la publier pour que d’autres utilisateurs y accèdent, vous devez déployer l’application web et mettre à jour le manifeste pour spécifier l’URL de l’application déployée.

Lorsque votre complément fonctionne comme vous le souhaitez, vous pouvez le publier directement via Visual Studio Code à l’aide de l’extension stockage Azure.

## <a name="using-visual-studio-code-to-publish"></a>Utilisation de Visual Studio Code pour publier

>[!NOTE]
> Ces étapes fonctionnent uniquement pour les projets créés avec le générateur Yeoman.

1. Ouvrez votre projet à partir de son dossier racine dans Visual Studio Code (VS Code).
2. À partir de la vue Extensions dans VS Code, recherchez l’extension stockage Azure et installez-la.
3. Une fois installée, une icône Azure est ajoutée à la barre d’activités. Sélectionnez-la pour accéder à l’extension. Si votre barre d’activité est masquée, vous ne pourrez pas accéder à l’extension. Affichez la barre d’activité en sélectionnant **Afficher > Apparence > Barre d’activité**.
4. Exécutez l’extension et **sélectionnez Se connecter à Azure** pour vous connecter à votre compte Azure. Si vous n’avez pas encore de compte Azure, créez-en un en sélectionnant **Créer un compte Azure**. Suivez les étapes fournies pour configurer votre compte.
5. Une fois connecté, vos comptes de stockage Azure s’affichent dans l’extension. Si vous n’avez pas encore de compte de stockage, créez-en un à l’aide de l’option **Créer un compte de stockage** dans la palette de commandes. Nommez votre compte de stockage sous un nom global unique, en utilisant uniquement « a-z » et « 0-9 ». Notez que par défaut, cela crée un compte de stockage et un groupe de ressources portant le même nom. Il place automatiquement le compte de stockage dans la région USA Ouest. Cela peut être ajusté en ligne via [votre compte Azure](https://portal.azure.com/).
6. Sélectionnez et maintenez enfoncé (cliquez avec le bouton droit) votre compte de stockage, puis choisissez **Configurer le site web statique**. Vous serez invité à entrer le nom du document d’index et le nom du document 404. Remplacez le nom du document d’index par défaut `index.html` **`taskpane.html`** par . Vous pouvez également modifier le nom du document 404, mais ce n’est pas obligatoire.
7. Sélectionnez et maintenez votre stockage enfoncé (cliquez avec le bouton droit) à nouveau, puis **choisissez Parcourir le site web statique**. Dans la fenêtre du navigateur qui s’ouvre, copiez l’URL du site web.
8. Dans VS Code, ouvrez le fichier manifeste de votre projet (`manifest.xml`) et modifiez toute référence à votre URL localhost (par exemple `https://localhost:3000`) à l’URL que vous avez copiée. Ce point de terminaison est l’URL du site web statique pour votre compte de stockage nouvellement créé. Enregistrez les modifications apportées à votre fichier manifeste.
9. Ouvrez une invite de ligne de commande et accédez au répertoire racine de votre projet de complément. Exécutez ensuite la commande suivante pour préparer tous les fichiers pour le déploiement de production.

    ```command&nbsp;line
    npm run build
    ```

    Une fois la build terminée, le dossier **dist** dans le répertoire racine de votre projet de complément contient les fichiers que vous allez déployer dans les étapes suivantes.

10. Pour déployer, sélectionnez проводник, sélectionnez et maintenez enfoncé (cliquez avec le bouton droit) votre dossier **dist**, puis **choisissez Déployer sur le site web statique via Stockage Azure**. Lorsque vous y êtes invité, sélectionnez le compte de stockage que vous avez créé précédemment.

    :::image type="content" source="../images/deploy-to-static-website.png" alt-text="Sélectionnez le dossier dist, cliquez avec le bouton droit et choisissez Déployer sur un site web statique via Stockage Azure.":::

11. Une fois le déploiement terminé, cliquez avec le bouton droit sur le compte de stockage que vous avez créé précédemment, puis **choisissez Parcourir le site web statique**. Le site web statique s’ouvre et affiche le volet Office.

## <a name="deploy-custom-functions-for-excel"></a>Déployer des fonctions personnalisées pour Excel

Si votre complément a des fonctions personnalisées, il existe quelques étapes supplémentaires pour les activer sur le compte de stockage Azure. Tout d’abord, activez CORS pour qu’Office puisse accéder au fichier functions.json.

1. Cliquez avec le bouton droit sur le compte de stockage Azure et choisissez **Ouvrir dans le portail**.
1. Dans le groupe Paramètres, choisissez **Partage des ressources (CORS).** Vous pouvez également utiliser la zone de recherche pour trouver cette option.
1. Créez une règle CORS avec les paramètres suivants.

    |Propriété        |Valeur                        |
    |----------------|-----------------------------|
    |Origines autorisées | \*                          |
    |Méthodes autorisées | GET                         |
    |En-têtes autorisés | \*                          |
    |En-têtes exposés | Access-Control-Allow-Origin |
    |Âge maximal         | 200                          |

1. Cliquez sur **Enregistrer**.

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

1. Pour déployer, dans le **проводник**, sélectionnez et maintenez enfoncé (ou cliquez avec le bouton droit) le dossier **dist**, puis **choisissez Déployer sur le site web statique via Stockage Azure**. Lorsque vous y êtes invité, sélectionnez le compte de stockage que vous avez créé précédemment. Si vous avez déjà déployé le dossier **dist** , vous êtes invité à remplacer les fichiers du stockage Azure par les dernières modifications.

## <a name="see-also"></a>Voir aussi

- [Développement d’un complément Office avec Visual Studio Code](../develop/develop-add-ins-vscode.md)
- [Déploiement et publication de votre complément Office](../publish/publish.md)
- [Prise en charge du partage de ressources cross-origin (CORS) pour Stockage Azure](/rest/api/storageservices/cross-origin-resource-sharing--cors--support-for-the-azure-storage-services)
