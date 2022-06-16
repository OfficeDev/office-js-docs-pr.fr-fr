---
title: Créer votre premier complément du volet des tâches de Project
description: Découvrez comment créer un complément simple de volet des tâches Project à l’aide de l’API JavaScript pour Office.
ms.date: 06/10/2022
ms.prod: project
ms.localizationpriority: high
ms.openlocfilehash: 486752d53b49d305ed4cc52c07fd8043264f1e58
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091040"
---
# <a name="build-your-first-project-task-pane-add-in"></a>Créer votre premier complément du volet des tâches de Project

Cet article décrit comment créer un complément du volet des tâches de Project.

## <a name="prerequisites"></a>Conditions préalables

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Project 2016 ou version ultérieure pour Windows

## <a name="create-the-add-in"></a>Créer le complément

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Sélectionnez un type de projet :** `Office Add-in Task Pane project`
- **Sélectionnez un type de script :** `Javascript`
- **Comment souhaitez-vous nommer votre complément ?** `My Office Add-in`
- **Quelle application client Office voulez-vous prendre en charge ?** `Project`

![Capture d’écran montrant les invites et réponses relatives au générateur Yeoman dans une interface de ligne de commande.](../images/yo-office-project.png)

Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Explorer le projet

Le projet de complément que vous avez créé à l’aide du générateur Yeoman contient un exemple de code pour un complément de volet de tâches très simple.

- Le fichier **./manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.
- Le fichier **./src/taskpane/taskpane.html** contient les balises HTML du volet Office.
- Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.
- Le fichier **./src/taskpane/taskpane.js** contient le code d’API JavaScript pour Office qui facilite l’interaction entre le volet des tâches et l’application cliente Office.

## <a name="update-the-code"></a>Mettre à jour le code

Ouvrez le fichier **./src/taskpane/taskpane.js** dans votre éditeur de code et ajoutez le code suivant à la fonction `run`. Ce code utilise l’API JavaScript Office pour définir le champ `Name` et le champ `Notes` de la tâche sélectionnée.

```js
var taskGuid;

// Get the GUID of the selected task
Office.context.document.getSelectedTaskAsync(
    function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            taskGuid = result.value;

            // Set the specified fields for the selected task.
            var targetFields = [Office.ProjectTaskFields.Name, Office.ProjectTaskFields.Notes];
            var fieldValues = ['New task name', 'Notes for the task.'];

            // Set the field value. If the call is successful, set the next field.
            for (var i = 0; i < targetFields.length; i++) {
                Office.context.document.setTaskFieldAsync(
                    taskGuid,
                    targetFields[i],
                    fieldValues[i],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            i++;
                        }
                        else {
                            var err = result.error;
                            console.log(err.name + ' ' + err.code + ' ' + err.message);
                        }
                    }
                );
            }
        } else {
            var err = result.error;
            console.log(err.name + ' ' + err.code + ' ' + err.message);
        }
    }
);
```

## <a name="try-it-out"></a>Essayez

1. Accédez au dossier racine du projet.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. Démarrez le serveur web local.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    Exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre.

    ```command&nbsp;line
    npm run dev-server
    ```

1. Dans Project, créez un plan de projet simple.

1. Chargez votre complément dans Project en suivant les instructions fournies dans [Chargement de versions test de compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

1. Sélectionnez une seule tâche dans le projet.

1. Au bas du volet des tâches, sélectionnez le lien **Exécuter** pour renommer la tâche sélectionnée et ajouter des notes à la tâche sélectionnée.

    ![Capture d'écran de l'application Project avec le complément du volet des tâches chargé.](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé un complément de volet de tâches Project ! À présent, découvrez les fonctionnalités d’un complément Project et explorez des scénarios courants.

> [!div class="nextstepaction"]
> [Compléments Project](../project/project-add-ins.md)

## <a name="see-also"></a>Voir aussi

- [Développement de compléments Office](../develop/develop-overview.md)
- [Concepts de base pour les compléments Office](../overview/core-concepts-office-add-ins.md)
- [Utilisation de Visual Studio Code pour publier](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)