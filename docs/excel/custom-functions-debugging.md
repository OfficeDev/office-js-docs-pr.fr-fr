---
title: Débogage des fonctions personnalisées sans interface utilisateur
description: Découvrez comment déboguer vos Excel personnalisées qui n’utilisent pas de volet de tâches.
ms.date: 01/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1c50c52faa8e825b7724284ed3e440ccb425b407
ms.sourcegitcommit: 4a7b9b9b359d51688752851bf3b41b36f95eea00
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/22/2022
ms.locfileid: "63711041"
---
# <a name="ui-less-custom-functions-debugging"></a>Débogage des fonctions personnalisées sans interface utilisateur

Cet article traite du débogage  uniquement pour les fonctions personnalisées qui n’utilisent pas de volet de tâches ou d’autres éléments d’interface utilisateur (fonctions personnalisées sans interface utilisateur).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

Sur Windows :

- [Excel déboeur de bureau et Visual Studio Code (VS Code)](#use-the-vs-code-debugger-for-excel-desktop)
- [Excel sur le Web et VS Code débogger](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [Excel sur le Web et les outils de navigateur](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [Ligne de commande](#use-the-command-line-tools-to-debug)

Sur Mac :

- [Excel sur le Web et les outils de navigateur](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [Ligne de commande](#use-the-command-line-tools-to-debug)

> [!NOTE]
> Par souci de simplicité, cet article présente le débogage dans le contexte de l’utilisation de Visual Studio Code pour modifier, exécuter des tâches et, dans certains cas, utiliser l’affichage débogage. Si vous utilisez un autre éditeur ou outil de ligne de commande, consultez les [instructions](#commands-for-building-and-running-your-add-in) de ligne de commande à la fin de cet article.

## <a name="requirements"></a>Conditions requises

Ce processus de débogage  fonctionne uniquement pour les fonctions personnalisées sans interface utilisateur, qui n’utilisent pas de volet de tâches ou d’autres éléments d’interface utilisateur. Pour créer une fonction personnalisée sans interface utilisateur, suivez les étapes du didacticiel Créer des fonctions [personnalisées dans Excel](../tutorials/excel-tutorial-create-custom-functions.md), puis supprimez tous les éléments du volet Des tâches et de l’interface utilisateur installés par le générateur [Yeoman](../develop/yeoman-generator-overview.md) pour les Office.

Notez que ce processus de débogage n’est pas compatible avec les projets de fonctions personnalisées à l’aide [d’un runtime partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>Utiliser le débo VS Code débogger pour Excel bureau

Vous pouvez utiliser VS Code pour déboguer des fonctions personnalisées sans interface utilisateur Office Excel sur le Bureau.

> [!NOTE]
> Le débogage du bureau pour Mac n’est pas disponible, mais peut être réalisé à l’aide des outils de navigateur et de la ligne de commande pour [déboguer Excel sur le Web](#use-the-command-line-tools-to-debug)).

### <a name="run-your-add-in-from-vs-code"></a>Exécuter votre VS Code

1. Ouvrez le dossier de projet racine de vos fonctions personnalisées [dans VS Code](https://code.visualstudio.com/).
1. Choisissez **Terminal > exécuter la tâche et** tapez ou sélectionnez **Regarder**. Cela surveillera et reconstruira les modifications apportées aux fichiers.
1. Choisissez **Terminal > exécuter la tâche** et tapez ou sélectionnez **Serveur dev**.

### <a name="start-the-vs-code-debugger"></a>Démarrer le débo VS Code débompeur

1. Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.
1. Dans le menu déroulant Exécuter, choisissez **Excel bureau (fonctions personnalisées).**
1. **Sélectionnez F5** (ou **exécutez -> démarrer** le débogage à partir du menu) pour commencer le débogage. Un nouveau Excel de travail s’ouvre avec votre add-in déjà chargé et prêt à l’emploi.

### <a name="start-debugging"></a>Démarrer le débogage

1. Dans VS Code, ouvrez votre fichier de script de code source (**functions.js** ou **functions.ts**).
2. [Définissez un point d’arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.
3. Dans le Excel, entrez une formule qui utilise votre fonction personnalisée.

À ce stade, l’exécution s’arrête sur la ligne de code où vous définissez le point d’arrêt. Vous pouvez maintenant vous servir de votre code, définir des montres et utiliser VS Code fonctionnalités de débogage dont vous avez besoin.

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a>Utilisez le débo VS Code débogger pour Excel dans Microsoft Edge

Vous pouvez utiliser VS Code pour déboguer des fonctions personnalisées sans interface utilisateur Excel sur le Microsoft Edge navigateur. Pour utiliser VS Code avec Microsoft Edge, vous devez installer [Microsoft Edge’extension DevTools pour Visual Studio Code](/microsoft-edge/visual-studio-code/microsoft-edge-devtools-extension).

### <a name="run-your-add-in-from-vs-code"></a>Exécuter votre VS Code

1. Ouvrez le dossier de projet racine de vos fonctions personnalisées [dans VS Code](https://code.visualstudio.com/).
2. Choisissez **Terminal > exécuter la tâche et** tapez ou sélectionnez **Regarder**. Cela surveillera et reconstruira les modifications apportées aux fichiers.
3. Choisissez **Terminal > exécuter la tâche** et tapez ou sélectionnez **Serveur dev**.

### <a name="start-the-vs-code-debugger"></a>Démarrer le débo VS Code débompeur

1. Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.
1. Dans les options Débogage, choisissez **Office Online (Edge Chromium).**
1. Ouvrez Excel dans le navigateur Microsoft Edge et créez un nouveau workbook.
1. Choisissez **Partager** dans le ruban et copiez le lien pour l’URL de ce nouveau workbook.
1. **Sélectionnez F5** (ou **exécutez > démarrer le** débogage à partir du menu) pour commencer le débogage. Une invite s’affiche, qui demande l’URL de votre document.
1. Collez l’URL de votre workbook et appuyez sur Entrée.

### <a name="sideload-your-add-in"></a>Charger une version test de votre complément

1. Sélectionnez **l’onglet** Insérer sur le ruban et, dans **la section Des** Office, **sélectionnez Les modules.**
2. Dans **la boîte Office** de dialogue Des Office, sélectionnez l’onglet MES **MODULES**. Sélectionnez Gérer mes **applications, puis** Télécharger Mon **add-in**.
  
    ![La boîte de dialogue Office des applications avec une zone de texte dans le coin supérieur droit de la lecture « Gérer mes applications » et une zone de texte en dessous avec l’option « Télécharger Mon add-in ».](../images/office-add-ins-my-account.png)

3. **Accédez** au fichier manifeste du add-in, puis **sélectionnez Télécharger**.
  
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

### <a name="set-breakpoints"></a>Définir des points d’arrêt

1. Dans VS Code, ouvrez votre fichier de script de code source (**functions.js** ou **functions.ts**).
2. [Définissez un point d’arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.
3. Dans le Excel, entrez une formule qui utilise votre fonction personnalisée.

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a>Utiliser les outils de développement du navigateur pour déboguer des fonctions personnalisées dans Excel sur le Web

Vous pouvez utiliser les outils de développement du navigateur pour déboguer des fonctions personnalisées sans interface utilisateur dans Excel sur le Web. Les étapes suivantes fonctionnent pour Windows macOS.

### <a name="run-your-add-in-from-visual-studio-code"></a>Exécuter votre Visual Studio Code

1. Ouvrez le dossier de projet racine de vos fonctions personnalisées [dans Visual Studio Code (VS Code)](https://code.visualstudio.com/).
2. Choisissez **Terminal > exécuter la tâche et** tapez ou sélectionnez **Regarder**. Cela surveillera et reconstruira les modifications apportées aux fichiers.
3. Choisissez **Terminal > exécuter la tâche** et tapez ou sélectionnez **Serveur dev**.

### <a name="sideload-your-add-in"></a>Charger une version test de votre complément

1. [Ouvrez Office sur le Web](https://office.live.com/).
2. Ouvrez un nouveau Excel de travail.
3. Ouvrez **l’onglet** Insérer sur le ruban et, dans **la section Des** Office, sélectionnez **Les modules.**
4. Dans **la boîte Office** de dialogue Des Office, sélectionnez l’onglet MES **MODULES**. Sélectionnez Gérer mes **applications, puis** Télécharger Mon **add-in**.
  
    ![La boîte de dialogue Office des applications avec une zone de texte dans le coin supérieur droit de la lecture « Gérer mes applications » et une zone de texte en dessous avec l’option « Télécharger Mon add-in ».](../images/office-add-ins-my-account.png)

5. **Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.
  
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

> [!NOTE]
> Une fois que vous avez chargé une version de version sideload dans le document, celui-ci reste chargé de nouveau à chaque ouverture du document.

### <a name="start-debugging"></a>Démarrer le débogage

1. Ouvrez les outils de développement dans le navigateur. Pour Chrome et la plupart des navigateurs F12 ouvrent les outils de développement.
2. Dans les outils de développement, ouvrez votre fichier de script de code source à l’aide de **Cmd+P** ou **Ctrl+P** (**functions.js** ou **functions.ts**).
3. [Définissez un point d’arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée. 

Si vous devez modifier le code, vous pouvez effectuer des modifications dans VS Code et enregistrer les modifications. Actualisez le navigateur pour voir les modifications chargées.

## <a name="use-the-command-line-tools-to-debug"></a>Utiliser les outils de ligne de commande pour déboguer

Si vous n’utilisez pas VS Code, vous pouvez utiliser la ligne de commande (par exemple, Bash ou PowerShell) pour exécuter votre module. Vous devez utiliser les outils de développement du navigateur pour déboguer votre code dans Excel sur le Web. Vous ne pouvez pas déboguer la version de bureau de Excel à l’aide de la ligne de commande.

1. À partir de la ligne de commande, exécutez `npm run watch` la commande pour observer et reconstruire lorsque des modifications de code se produisent.
2. Ouvrez une deuxième fenêtre de ligne de commande (la première sera bloquée lors de l’exécution de l’observation).)

3. Si vous souhaitez démarrer votre application dans la version de bureau de Excel, exécutez la commande suivante.
  
    `npm run start:desktop`
  
    Ou si vous préférez démarrer votre Excel sur le Web exécutez la commande suivante.
  
    `npm run start:web -- --document {url}`(où se `{url}` trouve l’URL d’un fichier Excel sur OneDrive ou SharePoint)
  
    Si votre add-in ne charge pas une version de version sideload dans le document, suivez les étapes du chargement de version [sideload](#sideload-your-add-in) de votre add-in. Ensuite, continuez jusqu’à la section suivante pour démarrer le débogage.
  
4. Ouvrez les outils de développement dans le navigateur. Pour Chrome et la plupart des navigateurs F12 ouvrent les outils de développement.
5. Dans les outils de développement, ouvrez votre fichier de script de code source (**functions.js** ou **functions.ts**). Votre code de fonctions personnalisées peut se trouver à la fin du fichier.
6. Dans le code source de la fonction personnalisée, appliquez un point d’arrêt en sélectionnant une ligne de code.

Si vous devez modifier le code, vous pouvez effectuer des modifications dans Visual Studio et enregistrer les modifications. Actualisez le navigateur pour voir les modifications chargées.

### <a name="commands-for-building-and-running-your-add-in"></a>Commandes de création et d’exécution de votre add-in

Plusieurs tâches de build sont disponibles.

- `npm run watch`: se construit pour le développement et se reconstruit automatiquement lorsqu’un fichier source est enregistré
- `npm run build-dev`: builds pour le développement une fois
- `npm run build`: builds pour la production
- `npm run dev-server`: exécute le serveur web utilisé pour le développement

Vous pouvez utiliser les tâches suivantes pour démarrer le débogage sur un ordinateur de bureau ou en ligne.

- `npm run start:desktop`: démarre Excel sur ordinateur de bureau et charge une version de version de chargement de version de votre application.
- `npm run start:web -- --document {url}`(où se `{url}` trouve l’URL d’un fichier Excel sur OneDrive ou SharePoint) : démarre Excel sur le Web et charge une version de votre module.
- `npm run stop`: arrête Excel et le débogage.

## <a name="next-steps"></a>Prochaines étapes

Découvrez les [pratiques d’authentification pour les fonctions personnalisées sans interface utilisateur](custom-functions-authentication.md).

## <a name="see-also"></a>Voir aussi

* [Résolution des problèmes des fonctions personnalisées](custom-functions-troubleshooting.md)
* [Gestion des erreurs liées aux fonctions personnalisées dans Excel](custom-functions-errors.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
