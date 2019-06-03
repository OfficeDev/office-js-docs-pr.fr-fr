---
title: Créer votre premier complément du volet des tâches de Project
description: ''
ms.date: 05/08/2019
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: 7a7c907eeeb85b2a686c49ebba0558f4ec20568d
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/30/2019
ms.locfileid: "34589208"
---
# <a name="build-your-first-project-task-pane-add-in"></a>Créer votre premier complément du volet des tâches de Project

Cet article décrit comment créer un complément du volet des tâches de Project.

## <a name="prerequisites"></a>Conditions préalables

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Project 2016 ou version ultérieure pour Windows

## <a name="create-the-add-in"></a>Créer le complément

1. Utilisez le générateur Yeoman afin de créer un projet de complément Project. Exécutez la commande suivante, puis répondez aux invites comme suit :

    ```command&nbsp;line
    yo office
    ```

    - **Sélectionnez un type de projet :** `Office Add-in Task Pane project`
    - **Sélectionnez un type de script :** `Javascript`
    - **Comment souhaitez-vous nommer votre complément ?** `My Office Add-in`
    - **Quelle application client Office voulez-vous prendre en charge ?** `Project`

    ![Capture d’écran des invites et des réponses relatives au générateur Yeoman](../images/yo-office-project.png)
    
    Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants de nœud de la prise en charge.
    
2. Accédez au dossier racine du projet.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

## <a name="explore-the-project"></a>Explorer le projet

Le projet de complément que vous avez créé à l’aide du générateur Yeoman contient un exemple de code pour un complément de volet de tâches très simple. 

- Le fichier **./manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.
- Le fichier **./src/taskpane/taskpane.html** contient les balises HTML du volet Office.
- Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.
- Le fichier **./src/taskpane/taskpane.js** contient le code de l’API JavaScript pour Office qui facilite l’interaction entre le volet Office et l’application hôte Office.

## <a name="update-the-code"></a>Mettre à jour le code

Dans votre éditeur de code, ouvrez le fichier **./src/taskpane/taskpane.js** et ajoutez le code suivant à la fonction **run**. Ce code utilise l’API JavaScript Office pour définir le champ `Name` et le champ `Notes` de la tâche sélectionnée.

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

## <a name="try-it-out"></a>Try it out

> [!NOTE]
> Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez. Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman.

1. Exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre.

    ```command&nbsp;line
    npm start
    ```

2. Dans Project, créez un plan de projet simple.

3. Chargez votre complément dans Project en suivant les instructions fournies dans [Chargement de versions test de compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

4. Sélectionnez une seule tâche dans le projet.

5. Au bas du volet des tâches, sélectionnez le lien **Exécuter** pour renommer la tâche sélectionnée et ajouter des notes à la tâche sélectionnée.

    ![Capture d’écran de l’application Project avec le complément du volet des tâches chargé](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé un complément du volet des tâches de Project ! Ensuite, découvrez les fonctionnalités d’un complément Project et explorez des scénarios plus courants.

> [!div class="nextstepaction"]
> [Compléments Project](../project/project-add-ins.md)

