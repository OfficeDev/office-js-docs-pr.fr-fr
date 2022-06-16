---
title: Créer votre premier complément du volet Office de OneNote
description: Découvrez comment créer un complément simple de volet des tâches OneNote simple à l’aide de l’API JavaScript pour Office.
ms.date: 06/10/2022
ms.prod: onenote
ms.localizationpriority: high
ms.openlocfilehash: 9b5f4dd941ed8cc107bee04bc67a368520439948
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66090851"
---
# <a name="build-your-first-onenote-task-pane-add-in"></a>Créer votre premier complément du volet Office de OneNote

Cet article décrit comment créer un complément du volet Office de OneNote.

## <a name="prerequisites"></a>Conditions préalables

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a>Création du projet de complément

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Sélectionnez un type de projet :** `Office Add-in Task Pane project`
- **Sélectionnez un type de script :** `Javascript`
- **Comment souhaitez-vous nommer votre complément ?** `My Office Add-in`
- **Quelle application client Office voulez-vous prendre en charge ?** `OneNote`

![Capture d’écran montrant les invites et réponses relatives au générateur Yeoman dans une interface de ligne de commande.](../images/yo-office-onenote.png)

Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Explorer le projet

Le projet de complément que vous avez créé à l’aide du générateur Yeoman contient un exemple de code pour un complément de volet de tâches très simple.

- Le fichier **./manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.
- Le fichier **./src/taskpane/taskpane.html** contient les balises HTML du volet Office.
- Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.
- Le fichier **./src/taskpane/taskpane.js** contient le code d’API JavaScript pour Office qui facilite l’interaction entre le volet des tâches et l’application cliente Office.

## <a name="update-the-code"></a>Mettre à jour le code

Ouvrez le fichier **./src/taskpane/taskpane.js** dans l’éditeur de code et ajoutez le code suivant à la fonction `run`. Ce code utilise l’API JavaScript OneNote pour définir le titre de la page et ajouter un plan au corps de celle-ci.

```js
try {
    await OneNote.run(async (context) => {

        // Get the current page.
        var page = context.application.getActivePage();

        // Queue a command to set the page title.
        page.title = "Hello World";

        // Queue a command to add an outline to the page.
        var html = "<p><ol><li>Item #1</li><li>Item #2</li></ol></p>";
        page.addOutline(40, 90, html);

        // Run the queued commands.
        await context.sync();
    });
} catch (error) {
    console.log("Error: " + error);
}
```

## <a name="try-it-out"></a>Essayez

1. Accédez au dossier racine du projet.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. Démarrez le serveur web local. Exécutez la commande suivante dans le répertoire racine de votre projet.

    ```command&nbsp;line
    npm run dev-server
    ```

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

1. Dans [OneNote sur le web](https://www.onenote.com/notebooks), ouvrez un bloc-notes, puis créez une page.

1. Choisissez **Insertion > Compléments Office** pour ouvrir la boîte de dialogue Compléments Office.

    - Si vous êtes connecté avec votre compte de consommateur, sélectionnez l’onglet **MES COMPLÉMENTS**, puis choisissez **Télécharger mon complément**.

    - Si vous êtes connecté avec votre compte professionnel ou scolaire, sélectionnez l’onglet **MON ORGANISATION**, puis choisissez **Télécharger mon complément**.

    L’image suivante montre l’onglet **MES COMPLÉMENTS** pour les blocs-notes de consommateurs.

    ![Capture d’écran de la boîte de dialogue Compléments Office affichant l’onglet MES COMPLÉMENTS.](../images/onenote-office-add-ins-dialog.png)

1. Dans la boîte de dialogue Télécharger le complément, accédez à **manifest.xml** dans le dossier de projet, puis choisissez **Télécharger**.

1. Dans l’onglet **Accueil**, choisissez le bouton **Afficher le volet de tâches** du ruban. Le volet Office du complément s’ouvre dans un iFrame à côté de la page OneNote.

1. Au bas du volet Office, sélectionnez le lien **Exécuter** pour définir le titre de la page et ajouter un plan au corps de celle-ci.

    ![Capture d’écran illustrant le complément créé à partir de cette procédure : bouton Afficher le ruban du volet Office et le volet Office dans OneNote.](../images/onenote-first-add-in-4.png)

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé un complément de volet office OneNote ! Découvrez ensuite les concepts fondamentaux de la création de compléments OneNote.

> [!div class="nextstepaction"]
> [Vue d’ensemble de la programmation de l’API JavaScript de OneNote](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
- [Développement de compléments Office](../develop/develop-overview.md)
- [Vue d’ensemble de la programmation de l’API JavaScript de OneNote](../onenote/onenote-add-ins-programming-overview.md)
- [Référence de l’API JavaScript de OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)
- [Exemple de grille d’évaluation](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Utilisation de Visual Studio Code pour publier](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)