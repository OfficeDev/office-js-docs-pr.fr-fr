---
title: Créer un complément Office autonome à partir de votre code Script Lab
description: Découvrez comment déplacer votre extrait de code de Script Lab vers un projet Yo Office
ms.topic: how-to
ms.date: 04/07/2022
ms.localizationpriority: high
ms.openlocfilehash: 038d25610e5ef5cc3e4cdbedb2d2a184294c673e
ms.sourcegitcommit: 5ef2c3ed9eb92b56e36c6de77372d3043ad5b021
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/15/2022
ms.locfileid: "64863297"
---
# <a name="create-a-standalone-office-add-in-from-your-script-lab-code"></a>Créer un complément Office autonome à partir de votre code Script Lab

Si vous avez créé un extrait de code dans Script Lab, vous pouvez le transformer en complément autonome. Vous pouvez copier le code de Script Lab dans un projet généré par le [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md) (également appelé « Yo Office »). Vous pouvez ensuite continuer à développer le code en tant que complément que vous pouvez éventuellement déployer sur d’autres utilisateurs.

Les étapes décrites dans cet article font référence à [Visual Studio Code](https://code.visualstudio.com/), mais vous pouvez utiliser n’importe quel éditeur de code de votre choix.

## <a name="create-a-new-yo-office-project"></a>Créer un projet Yo Office

Vous devez créer le projet de complément autonome qui sera le nouvel emplacement de développement pour votre code d’extrait de code.

Exécutez la commande `yo office --projectType taskpane --ts true --host <host> --name "basic-sample"`, où `<host>` est l’une des valeurs suivantes.

- excel
- outlook
- powerpoint
- word

> [!IMPORTANT]
> La valeur de l’argument `--name` doit être entre guillemets doubles, même si elle n’a pas d’espace.

La commande précédente crée un dossier de projet nommé **exemple de base**. Il est configuré pour s’exécuter dans l’hôte que vous avez spécifié et utilise TypeScript. Script Lab utilise TypeScript par défaut, mais la plupart des extraits de code sont JavaScript. Vous pouvez générer un projet JavaScript Yo Office si vous préférez, mais assurez-vous simplement que le code que vous copiez est JavaScript.

## <a name="open-the-snippet-in-script-lab"></a>Ouvrir l’extrait de code dans script Lab

Utilisez un extrait de code existant dans Script Lab pour apprendre à copier un extrait de code dans un projet généré par Yo Office.

1. Ouvrez Office (Word, Excel, PowerPoint ou Outlook), puis ouvrez Script Lab.
1. Sélectionnez **Script Lab** >  **Code**. Si vous travaillez dans Outlook, ouvrez un e-mail pour voir Script Lab sur le ruban.
1. Dans le volet Office Script Lab, choisissez **Exemples**. Sélectionnez ensuite un exemple de base en fonction de l’hôte Office dans lequel vous travaillez.
    - Pour Excel ou Word, choisissez l’exemple **Appel d’API de base (TypeScript)**.
    - Pour Outlook, choisissez l’exemple **Utiliser les paramètres de complément** .
    - Pour PowerPoint, choisissez l’exemple **appel d’API de base (Ofice 2013)**.

## <a name="copy-snippet-code-to-visual-studio-code"></a>Copier le code d’extrait de code dans Visual Studio code

Vous pouvez maintenant copier le code de l’extrait de code vers le projet Yo Office dans VS Code.

- Dans VS Code, ouvrez le projet **exemple de base**.

Dans les étapes suivantes, vous allez copier le code à partir de plusieurs onglets dans Script Lab.

:::image type="content" source="../images/script-lab-script-tabs.png" alt-text="Capture d’écran des onglets dans Script Lab.":::

### <a name="copy-task-pane-code"></a>Copier le code du volet Office

1. Dans VS Code, ouvrez le fichier **/src/taskpane/taskpane.ts**. Si vous utilisez un projet JavaScript, le nom de fichier est **taskpane.js**.
1. Dans Script Lab, sélectionnez l’onglet **Script** .
1. Copiez tout le code dans l’onglet **Script** dans le Presse-papiers. Remplacez l’intégralité du contenu de **taskpane.ts** (ou **taskpane.js** pour javaScript) par le code que vous avez copié.

### <a name="copy-task-pane-html"></a>Copier le code HTML du volet Office

1. Dans VS Code, ouvrez le fichier **/src/taskpane/taskpane.html**.
1. Dans Script Lab, sélectionnez l’onglet **HTML**.
1. Copiez tout le code HTML de l’onglet **HTML** dans le Presse-papiers. Remplacez tout le code HTML à l’intérieur de la balise `<body>` par le code HTML que vous avez copié.

### <a name="copy-task-pane-css"></a>Copier le CSS du volet Office

1. Dans VS Code, ouvrez le fichier **/src/taskpane/taskpane.css**.
1. Dans Script Lab, sélectionnez l’onglet **CSS**.
1. Copiez tous les CSS de l’onglet **CSS** dans le Presse-papiers. Remplacez l’intégralité du contenu de **taskpane.css** par le CSS que vous avez copié.
1. Enregistrez toutes les modifications apportées aux fichiers que vous avez mis à jour lors des étapes précédentes.

## <a name="add-jquery-support"></a>Ajouter la prise en charge de jQuery

Script Lab utilise jQuery dans les extraits de code. Vous devez ajouter cette dépendance au projet Yo Office pour exécuter le code correctement.

1. Ouvrez le fichier **taskpane.html** et ajoutez la balise de script suivante à la section`<head>`.

    ```html
     <script src="https://ajax.aspnetcdn.com/ajax/jquery/jquery-3.3.1.js"></script>
    ```

    > [!NOTE]
    > La version spécifique de jQuery peut varier. Vous pouvez déterminer la version que Script Lab utilise en choisissant l’onglet **Bibliothèques**.

1. Ouvrez un terminal dans VS Code et entrez les commandes suivantes.

    ```command&nbsp;line
    npm install --save-dev jquery@3.1.1
    npm install --save-dev @types/jquery@3.3.1
    ```

Si vous avez créé un extrait de code qui a des dépendances de bibliothèque supplémentaires, veillez à les ajouter au projet Yo Office. Recherchez la liste de toutes les dépendances de bibliothèque sous l’onglet **Bibliothèques** dans Script Lab.

## <a name="handle-initialization"></a>Gérer l’initialisation

Script Lab gère automatiquement l’initialisation`Office.onReady`. Vous devez modifier le code pour fournir votre propre gestionnaire de `Office.onReady` .

1. Ouvrez le fichier **taskpane.ts** (ou **taskpane.js** pour JavaScript).
1. Pour Excel ou Word, remplacez :

    ```typescript
    $("#run").click(() => tryCatch(run));
    ```

    avec :

    ```typescript
    Office.onReady(function () {
      // Office is ready
      $(document).ready(function () {
        // The document is ready
        $("#run").click(() => tryCatch(run));
      });
    });
    ```

1. Pour Outlook, remplacez :

    ```typescript
    $("#get").click(get);
    $("#set").click(set);
    $("#save").click(save);
    ```

    avec :

    ```typescript
    Office.onReady(function () {
      // Office is ready
      $(document).ready(function () {
        // The document is ready
        $("#get").click(get);
        $("#set").click(set);
        $("#save").click(save);
      });
    });
    ```

1. Pour PowerPoint, remplacez :

    ```typescript
    $("#run").click(run);
    ```

    avec :

    ```typescript
    Office.onReady(function () {
      // Office is ready
      $(document).ready(function () {
        // The document is ready
        $("#run").click(run);
      });
    });
    ```

1. Enregistrez le fichier.

## <a name="custom-functions"></a>Fonctions personnalisées

Si votre extrait de code utilise des fonctions personnalisées, vous devez utiliser le modèle de fonctions personnalisées Yo Office. Pour transformer des fonctions personnalisées en complément autonome, procédez comme suit.

1. Exécutez la commande `yo office --projectType excel-functions --ts true --name "functions-sample"`.

    > [!IMPORTANT]
    > La valeur de l’argument `--name` doit être entre guillemets doubles, même si elle n’a pas d’espace.

1. Ouvrez Excel, puis ouvrez Script Lab.
1. Sélectionnez **Script Lab** >  **Code**.
1. Dans le volet office Script Lab, choisissez **Exemples**, puis choisissez l’exemple **Fonction personnalisée de base**.
1. Ouvrez le fichier **/src/functions/functions.ts** . Si vous utilisez un projet JavaScript, le nom de fichier est **functions.js**.
1. Dans Script Lab, sélectionnez l’onglet **Script** .
1. Copiez tout le code dans l’onglet **Script** dans le Presse-papiers. Collez le code en haut de **functions.ts** (ou **functions.js** pour javaScript) avec le code que vous avez copié.
1. Enregistrez le fichier.

## <a name="test-the-standalone-add-in"></a>Tester le complément autonome

Une fois toutes les étapes terminées, exécutez et testez votre complément autonome. Exécutez la commande suivante pour commencer.

```command&nbsp;line
npm start
```

Office démarre et vous pouvez ouvrir le volet Office de votre complément à partir du ruban. Félicitations ! Vous pouvez maintenant continuer à créer votre complément en tant que projet autonome.

## <a name="console-logging"></a>Journalisation de la console

De nombreux extraits de code dans script lab écrivent la sortie dans une section de console en bas du volet Office. Le projet Yo Office n’a pas de section de console. Toutes les instructions `console.log*` écrivent dans la console de débogage par défaut (par exemple, les outils de développement de votre navigateur). Si vous souhaitez que la sortie accède à votre volet Office, vous devez mettre à jour le code.
