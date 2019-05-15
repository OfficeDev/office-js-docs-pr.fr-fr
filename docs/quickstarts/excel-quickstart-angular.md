---
title: Créer un complément de volet de tâches Excel à l’aide d’Angular
description: ''
ms.date: 05/02/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 66c85ba9914b783295e9ed2143dc9ce107f64c4c
ms.sourcegitcommit: 47b792755e655043d3db2f1fdb9a1eeb7453c636
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33619907"
---
# <a name="build-an-excel-task-pane-add-in-using-angular"></a>Créer un complément de volet de tâches Excel à l’aide d’Angular

Cet article décrit le processus de création d’un complément de volet de tâches Excel à l’aide d’Angular et de l’API JavaScript pour Excel.

## <a name="prerequisites"></a>Conditions préalables

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a>Création du projet de complément

1. Utilisez le générateur Yeoman pour créer un projet de complément Excel. Exécutez la commande suivante, puis répondez aux invites comme suit :

    ```command&nbsp;line
    yo office
    ```

    - **Sélectionnez un type de projet :** `Office Add-in Task Pane project using Angular framework`
    - **Sélectionnez un type de script :** `TypeScript`
    - **Comment souhaitez-vous nommer votre complément ?** `My Office Add-in`
    - **Quelle application client Office voulez-vous prendre en charge ?** `Excel`

    ![Générateur Yeoman](../images/yo-office-excel-angular-2.png)

    Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants de nœud de la prise en charge.

2. Accédez au dossier racine du projet.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```
## <a name="explore-the-project"></a>Explorer le projet

Le projet de complément que vous avez créé à l’aide du générateur Yeoman contient un exemple de code pour un complément de volet de tâches très simple. Pour explorer les composants clés de votre projet de complément, ouvrez le projet dans votre éditeur de code et passez en revue les fichiers répertoriés ci-dessous. Lorsque vous êtes prêt à tester votre complément, passez à la section suivante.

- Le fichier **manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.
- Le fichier **./src/taskpane/app/app.component.html** contient les balises HTML du volet Office.
- Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.
- Le fichier **./src/taskpane/app/app.component.ts** contient le code d’API JavaScript pour Office qui facilite l’interaction entre le volet Office et Excel.

## <a name="try-it-out"></a>Essayez !

1. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

2. Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.

    ![Bouton Complément Excel](../images/excel-quickstart-addin-3b.png)

3. Sélectionnez une plage de cellules dans la feuille de calcul.

4. En bas du volet Office, cliquez sélectionnez le lien **Exécuter** pour définir la couleur de la plage sélectionnée sur jaune.

    ![Complément Excel](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé un complément de volet de tâches Excel à l’aide d’Angular ! Ensuite, découvrez les fonctionnalités d’un complément Excel et créez-en un plus complexe en suivant le didacticiel sur les compléments Excel.

> [!div class="nextstepaction"]
> [Didacticiel sur les compléments Excel](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>Voir aussi

* [Didacticiel sur les compléments Excel](../tutorials/excel-tutorial-create-table.md)
* [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](../excel/excel-add-ins-core-concepts.md)
* [Exemples de code pour les compléments Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Référence de l’API JavaScript pour Excel](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
