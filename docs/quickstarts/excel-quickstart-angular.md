---
title: Utiliser Angular pour créer un complément de volet Office Excel
description: Découvrez comment créer un complément de volet des tâches Excel simple à l’aide de l’API JavaScript et d’Angular pour Office.
ms.date: 06/10/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 372649188d8f617f65e0c2eddc4d758047b1a2cc
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091152"
---
# <a name="use-angular-to-build-an-excel-task-pane-add-in"></a>Utiliser Angular pour créer un complément de volet Office Excel

Cet article décrit le processus de création d’un complément de volet de tâches Excel à l’aide d’Angular et de l’API JavaScript pour Excel.

## <a name="prerequisites"></a>Conditions préalables

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a>Création du projet de complément

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Sélectionnez un type de projet :** `Office Add-in Task Pane project using Angular framework`
- **Sélectionnez un type de script :** `TypeScript`
- **Comment souhaitez-vous nommer votre complément ?** `My Office Add-in`
- **Quelle application client Office voulez-vous prendre en charge ?** `Excel`

![Capture d’écran de l’interface de ligne de commande du générateur de compléments Yeoman Office, avec l’option type de projet réglée sur l’infrastructure Angular.](../images/yo-office-excel-angular-2.png)

Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Explorer le projet

Le projet de complément que vous avez créé avec le générateur Yeoman contient un exemple de code pour un complément de volet Office de base. Pour explorer les composants clés de votre projet de complément, ouvrez le projet dans votre éditeur de code et passez en revue les fichiers répertoriés ci-dessous. Lorsque vous êtes prêt à tester votre complément, passez à la section suivante.

- Le fichier **manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément. Pour en savoir plus sur le fichier **manifest.xml**, consultez [manifeste XML des compléments Office](../develop/add-in-manifests.md).
- Le fichier **./src/taskpane/app/app.component.html** contient les balises HTML du volet Office.
- Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.
- Le fichier **./src/taskpane/app/app.component.ts** contient le code d’API JavaScript pour Office qui facilite l’interaction entre le volet Office et Excel.

## <a name="try-it-out"></a>Essayez

1. Accédez au dossier racine du projet.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)]

1. Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.

    ![Capture d’écran du menu Accueil d’ Excel, avec le bouton Afficher le volet Office mis en évidence.](../images/excel-quickstart-addin-3b.png)

1. Sélectionnez une plage de cellules dans la feuille de calcul.

1. En bas du volet Office, cliquez sélectionnez le lien **Exécuter** pour définir la couleur de la plage sélectionnée sur jaune.

    ![Capture d’écran d’Excel, avec le volet Office du complément ouvert et le bouton Exécuter mis en surbrillance dans ce volet.](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé un complément du volet Office Excel à l’aide d’Angular ! Maintenant, apprenez-en davantage sur les fonctionnalités d’un complément Excel et créez un complément plus complexe en suivant le didacticiel sur les compléments Excel.

> [!div class="nextstepaction"]
> [Didacticiel sur les compléments Excel](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
- [Développement de compléments Office](../develop/develop-overview.md)
- [Modèle d’objet JavaScript Excel dans les compléments Office](../excel/excel-add-ins-core-concepts.md)
- [Exemples de code pour les compléments Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Référence de l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Utilisation de Visual Studio Code pour publier](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
