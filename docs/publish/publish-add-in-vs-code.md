---
title: Publier un complément à l’aide de Visual Studio code et Azure
description: Comment publier un complément à l’aide de Visual Studio code et d’Azure Active Directory
ms.date: 08/12/2020
localization_priority: Normal
ms.openlocfilehash: 3552e4eebacc84fc2b8e37782c97b4e03e96e508
ms.sourcegitcommit: 7faa0932b953a4983a80af70f49d116c3236d81a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/21/2020
ms.locfileid: "46845509"
---
# <a name="publish-an-add-in-developed-with-visual-studio-code"></a>Publier un complément développé avec Visual Studio Code

Cet article explique comment un complément Office que vous avez créé à l’aide du générateur Yeoman et développé avec [Visual Studio Code (VS Code)](https://code.visualstudio.com) ou un autre éditeur.

> [!NOTE]
> Pour plus d’informations sur la publication d’un complément Office que vous avez créé à l’aide de Visual Studio, voir [Publier votre complément à l’aide de Visual Studio](package-your-add-in-using-visual-studio.md).

## <a name="publishing-an-add-in-for-other-users-to-access"></a>Publication d’un complément pour accéder à d’autres utilisateurs

Un complément Office comprend une application Web et un fichier manifeste. L’application Web définit l’interface utilisateur et les fonctionnalités du complément, tandis que le manifeste spécifie l’emplacement de l’application Web et définit les paramètres et fonctionnalités du complément.

Pendant que vous développez, vous pouvez exécuter le complément sur votre serveur Web local ( `localhost` ). Lorsque vous êtes prêt à le publier pour permettre à d’autres utilisateurs d’y accéder, vous devez déployer l’application Web et mettre à jour le manifeste pour spécifier l’URL de l’application déployée.

Lorsque votre complément fonctionne comme vous le souhaitez, vous pouvez le publier directement par le biais de Visual Studio code à l’aide de l’extension de stockage Azure.

## <a name="using-visual-studio-code-to-publish"></a>Utilisation de Visual Studio code pour publier

>[!NOTE]
> Ces étapes fonctionnent uniquement pour les projets créés avec le générateur Yeoman.

1. Ouvrez votre projet à partir de son dossier racine dans Visual Studio code (VS code).
2. À partir de l’affichage extensions dans le code VS, recherchez l’extension de stockage Azure et installez-la.
3. Une fois installé, une icône Azure est ajoutée à la barre d’activité. Sélectionnez-le pour accéder à l’extension. Si votre barre d’activité est masquée, vous ne pourrez pas accéder à l’extension. Affichez la barre d’activité en sélectionnant **afficher > apparence > afficher la barre d’activité**.
4. Lorsque vous êtes dans l’extension, connectez-vous à votre compte Azure en sélectionnant **se connecter à Azure**. Vous pouvez également créer un compte Azure si vous n’en avez pas encore en sélectionnant **créer un compte Azure gratuit**. Suivez les étapes décrites pour configurer votre compte.
5. Une fois que vous êtes connecté à votre compte Azure, les comptes de stockage Azure s’affichent dans l’extension. Si vous n’avez pas encore de compte de stockage, vous devez en créer un à l’aide de l’option **créer un nouveau compte de stockage** . Nommez votre compte de stockage un nom unique au format global, en utilisant uniquement « a-z » et « 0-9 ». Notez que, par défaut, cela crée un compte de stockage et un groupe de ressources portant le même nom. Il place automatiquement le compte de stockage dans l’ouest des États-Unis. Cela peut être ajusté en ligne via [votre compte Azure](https://portal.azure.com/).
6. Sélectionnez un compte de stockage et maintenez-le enfoncé (clic droit), en sélectionnant **configurer le site Web statique**. Vous serez invité à entrer le nom du document d’index et le nom du document 404. Remplacez le nom par défaut du document d’index par `index.html` **`taskpane.html`** . Vous pouvez également modifier le nom du document 404, mais vous n’êtes pas obligé de le faire.
7. Sélectionnez une nouvelle fois (cliquez avec le bouton droit) sur votre stockage, en sélectionnant **Parcourir le site Web statique**. Dans la fenêtre du navigateur qui s’ouvre, copiez l’URL du site Web.
8. Dans le code VS, ouvrez le fichier manifeste de votre projet ( `manifest.xml` ) et modifiez toute référence à votre URL localhost (telle que `https://localhost:3000` ) à l’URL que vous avez copiée. Ce point de terminaison est l’URL du site Web statique pour le compte de stockage que vous venez de créer. Enregistrez les modifications apportées à votre fichier manifeste.
9. Ouvrez une invite de ligne de commande et accédez au répertoire racine de votre projet de complément. Exécutez ensuite la commande suivante pour préparer tous les fichiers pour le déploiement de production.

    ```command&nbsp;line
    npm run build
    ```

    Une fois la build terminée, le dossier **dist** dans le répertoire racine de votre projet de complément contient les fichiers que vous allez déployer dans les étapes suivantes.

10. Pour déployer, sélectionnez l’Explorateur de fichiers, sélectionnez et maintenez (clic droit) votre dossier **dist** , puis choisissez **déployer vers un site Web statique**. Lorsque vous y êtes invité, sélectionnez le compte de stockage que vous avez créé précédemment.

![Déploiement sur un site Web statique](../images/deploy-to-static-website.png)

11. Lorsque le déploiement est terminé, un message **de navigation sur le site Web** s’affiche, que vous pouvez sélectionner pour ouvrir le point de terminaison principal du code d’application déployé.

## <a name="see-also"></a>Voir aussi

- [Développement d’un complément Office avec Visual Studio Code](../develop/develop-add-ins-vscode.md)
- [Déploiement et publication de votre complément Office](../publish/publish.md)
