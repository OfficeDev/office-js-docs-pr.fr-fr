---
title: Utiliser React pour créer un complément du volet Office Excel
description: Découvrez comment créer un complément de volet des tâches Excel simple à l’aide de l’API JavaScript et de React pour Office.
ms.date: 08/04/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 74a28f3914ddbc54188d3b8baa33fc1faa7a30fe
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58939198"
---
# <a name="use-react-to-build-an-excel-task-pane-add-in"></a>Utiliser React pour créer un complément du volet Office Excel

Cet article décrit le processus de création d’un complément de volet de tâches Excel à l’aide de React et de l’API JavaScript pour Excel.

## <a name="prerequisites"></a>Conditions préalables

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a>Création du projet de complément

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Sélectionnez un type de projet :** `Office Add-in Task Pane project using React framework`
- **Sélectionnez un type de script :** `TypeScript`
- **Comment souhaitez-vous nommer votre complément ?** `My Office Add-in`
- **Quelle application client Office voulez-vous prendre en charge ?** `Excel`

![Capture d’écran de l’interface de ligne de commande du générateur de compléments Yeoman Office, avec le type de projet défini sur l’infrastructure React.](../images/yo-office-excel-react-2.png)

Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Explorer le projet

Le projet de complément que vous avez créé avec le générateur Yeoman contient un exemple de code pour un complément de volet Office de base. Pour explorer les composants clés de votre projet de complément, ouvrez le projet dans votre éditeur de code et passez en revue les fichiers répertoriés ci-dessous. Lorsque vous êtes prêt à tester votre complément, passez à la section suivante.

- Le fichier **manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément. Pour en savoir plus sur le fichier **manifest.xml**, consultez [manifeste XML des compléments Office](../develop/add-in-manifests.md).
- Le fichier **./src/taskpane/taskpane.html** définit l’infrastructure HTML du volet de tâches et les fichiers du dossier **./src/taskpane/components** définissent les différentes parties de l’interface utilisateur du volet de tâches.
- Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet de tâches.
- Le fichier **./src/taskpane/component/App.tsx** contient le code d’API JavaScript pour Office qui facilite l’interaction entre le volet de tâches et Excel.

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

Félicitations, vous avez créé un complément du volet Office Excel à l’aide de React ! Maintenant, apprenez-en davantage sur les fonctionnalités d’un complément Excel et créez un complément plus complexe en suivant le didacticiel sur les compléments Excel.

> [!div class="nextstepaction"]
> [Didacticiel sur les compléments Excel](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>Voir aussi

- [Didacticiel sur les compléments Excel](../tutorials/excel-tutorial.md)
- [Modèle d’objet JavaScript Excel dans les compléments Office](../excel/excel-add-ins-core-concepts.md)
- [Exemples de code pour les compléments Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Référence de l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)