---
title: Publier un add-in à l’aide de Visual Studio Code azure
description: Comment publier un add-in à l’aide Visual Studio Code et Azure Active Directory
ms.date: 08/12/2020
ms.localizationpriority: medium
ms.openlocfilehash: a97a7ca2298af003a089f884ee2c409f0362a098
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153428"
---
# <a name="publish-an-add-in-developed-with-visual-studio-code"></a>Publier un complément développé avec Visual Studio Code

Cet article explique comment un complément Office que vous avez créé à l’aide du générateur Yeoman et développé avec [Visual Studio Code (VS Code)](https://code.visualstudio.com) ou un autre éditeur.

> [!NOTE]
> Pour plus d’informations sur la publication d’un complément Office que vous avez créé à l’aide de Visual Studio, voir [Publier votre complément à l’aide de Visual Studio](package-your-add-in-using-visual-studio.md).

## <a name="publishing-an-add-in-for-other-users-to-access"></a>Publication d’un complément pour accéder à d’autres utilisateurs

Un complément Office se compose d’une application web et d’un fichier manifeste. L’application web définit l’interface utilisateur et les fonctionnalités du complément, tandis que le manifeste spécifie l’emplacement de l’application web et définit les paramètres et les fonctionnalités du complément.

Pendant le développement, vous pouvez exécuter le add-in sur votre serveur web local ( `localhost` ). Lorsque vous êtes prêt à la publier pour que d’autres utilisateurs y accèdent, vous devez déployer l’application web et mettre à jour le manifeste pour spécifier l’URL de l’application déployée.

Lorsque votre add-in fonctionne comme vous le souhaitez, vous pouvez le publier directement Visual Studio Code l’aide de l stockage Azure extension.

## <a name="using-visual-studio-code-to-publish"></a>Utilisation de Visual Studio Code publier

>[!NOTE]
> Ces étapes fonctionnent uniquement pour les projets créés avec le générateur Yeoman.

1. Ouvrez votre projet à partir de son dossier racine dans Visual Studio Code (VS Code).
2. Dans la vue Extensions de VS Code, recherchez l’extension stockage Azure et installez-la.
3. Une fois installée, une icône Azure est ajoutée à la barre d’activité. Sélectionnez-le pour accéder à l’extension. Si votre barre d’activité est masquée, vous ne pourrez pas accéder à l’extension. Affichez la barre d’activité en **sélectionnant Afficher > l'> Afficher la barre d’activité.**
4. Dans l’extension, connectez-vous à votre compte Azure en sélectionnant **Se connectez à Azure.** Vous pouvez également créer un compte Azure si vous n’en avez pas déjà un en sélectionnant **Créer un compte Azure gratuit.** Suivez les étapes fournies pour configurer votre compte.
5. Une fois que vous vous êtes inscrit à votre compte Azure, vos comptes de stockage Azure apparaissent dans l’extension. Si vous n’avez pas encore de compte de stockage, vous devez en créer un à l’aide de l’option Créer un compte **de** stockage. Nommez votre compte de stockage un nom global unique, en utilisant uniquement « a-z » et « 0-9 ». Notez que, par défaut, cela crée un compte de stockage et un groupe de ressources du même nom. Il place automatiquement le compte de stockage dans l’Ouest des États-Unis. Cela peut être ajusté en ligne via [votre compte Azure.](https://portal.azure.com/)
6. Sélectionnez et maintenez (cliquez avec le bouton droit) votre compte de stockage, en sélectionnant **Configurer le site web statique.** Vous serez invité à entrer le nom du document d’index et le nom du document 404. Modifiez le nom du document d’index de la valeur par `index.html` défaut à **`taskpane.html`** . Vous pouvez également décider de modifier le nom du document 404, mais ce n’est pas obligatoire.
7. Sélectionnez et maintenez (cliquez avec le bouton droit) votre stockage à nouveau, cette fois en choisissant Parcourir le site **web statique.** Dans la fenêtre du navigateur qui s’ouvre, copiez l’URL du site web.
8. Dans VS Code, ouvrez le fichier manifeste de votre projet () et modifiez toute référence à votre `manifest.xml` URL localhost (par exemple) à l’URL que vous avez `https://localhost:3000` copiée. Ce point de terminaison est l’URL du site web statique de votre compte de stockage nouvellement créé. Enregistrez les modifications apportées à votre fichier manifeste.
9. Ouvrez une invite de ligne de commande et accédez au répertoire racine de votre projet de add-in. Exécutez ensuite la commande suivante pour préparer tous les fichiers au déploiement de production.

    ```command&nbsp;line
    npm run build
    ```

    Une fois la build terminée, le dossier **dist** dans le répertoire racine de votre projet de complément contient les fichiers que vous allez déployer dans les étapes suivantes.

10. Pour déployer, sélectionnez l’Explorateur de fichiers, sélectionnez et maintenez (cliquez avec le bouton droit) votre dossier **dist,** puis choisissez **Déployer sur le site web statique.** Lorsque vous y êtes invité, sélectionnez le compte de stockage que vous avez créé précédemment.

![Déploiement sur un site web statique.](../images/deploy-to-static-website.png)

11. Une fois le déploiement terminé, un message **Parcourir** vers le site web s’affiche et vous pouvez choisir d’ouvrir le point de terminaison principal du code de l’application déployée.

## <a name="see-also"></a>Voir aussi

- [Développement d’un complément Office avec Visual Studio Code](../develop/develop-add-ins-vscode.md)
- [Déploiement et publication de votre complément Office](../publish/publish.md)
