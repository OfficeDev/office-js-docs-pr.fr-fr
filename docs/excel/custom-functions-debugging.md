---
ms.date: 04/12/2021
description: Découvrez comment déboguer vos fonctions personnalisées Excel qui n'utilisent pas de volet de tâches.
title: Débogage de fonctions personnalisées sans interface utilisateur
localization_priority: Normal
ms.openlocfilehash: c6954af4638ae416c789af339d35187467e37b7f
ms.sourcegitcommit: 78fb861afe7d7c3ee7fe3186150b3fed20994222
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/26/2021
ms.locfileid: "52024324"
---
# <a name="ui-less-custom-functions-debugging"></a>Débogage des fonctions personnalisées sans interface utilisateur

Cet article traite du  débogage uniquement pour les fonctions personnalisées qui n'utilisent pas de volet de tâches ou d'autres éléments d'interface utilisateur (fonctions personnalisées sans interface utilisateur). 

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

Sur Windows :
- [Débogger Excel Desktop and Visual Studio Code (VS Code)](#use-the-vs-code-debugger-for-excel-desktop)
- [Débogger Excel sur le web et VS Code](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [Outils Excel sur le web et navigateur](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [Ligne de commande](#use-the-command-line-tools-to-debug)

Sur Mac :
- [Outils Excel sur le web et navigateur](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [Ligne de commande](#use-the-command-line-tools-to-debug)

> [!NOTE]
> Par souci de simplicité, cet article présente le débogage dans le contexte de l'utilisation de Visual Studio Code pour modifier, exécuter des tâches et, dans certains cas, utiliser l'affichage débogage. Si vous utilisez un autre éditeur ou outil de ligne de commande, consultez les [instructions](#commands-for-building-and-running-your-add-in) de ligne de commande à la fin de cet article.

## <a name="requirements"></a>Configuration requise

Ce processus de  débogage fonctionne uniquement pour les fonctions personnalisées sans interface utilisateur, qui n'utilisent pas de volet de tâches ou d'autres éléments d'interface utilisateur. Une fonction personnalisée sans interface utilisateur peut être créée en suivant les [étapes](../tutorials/excel-tutorial-create-custom-functions.md) du didacticiel Créer des fonctions personnalisées dans Excel, puis en supprimant tous les éléments du volet Office et de l'interface utilisateur installés par le générateur Yeoman pour les [add-ins Office.](https://www.npmjs.com/package/generator-office)

Notez que ce processus de débogage n'est pas compatible avec les projets de fonctions personnalisées à l'aide [d'un runtime partagé.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>Utiliser le débogger VS Code pour Excel Desktop

Vous pouvez utiliser VS Code pour déboguer des fonctions personnalisées sans interface utilisateur dans Office Excel sur le Bureau.

> [!NOTE]
> Le débogage du bureau pour Mac n'est pas disponible, mais peut être réalisé à l'aide des outils de navigateur et de la ligne de commande pour [déboguer Excel sur le web).](#use-the-command-line-tools-to-debug)

### <a name="run-your-add-in-from-vs-code"></a>Exécuter votre add-in à partir de VS Code

1. Ouvrez le dossier de projet racine de vos fonctions personnalisées dans [VS Code.](https://code.visualstudio.com/)
2. Choose **Terminal > Run Task** and type or select **Watch**. Cela surveillera et reconstruira les modifications apportées aux fichiers.
3. Choose **Terminal > Run Task** and type or select **Dev Server**.

### <a name="start-the-vs-code-debugger"></a>Démarrer le débogger VS Code

4. Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.
5. Dans le menu déroulant Exécuter, choisissez **Excel Desktop (Fonctions personnalisées).**
6. Sélectionnez **F5** (ou **exécutez -> démarrer le** débogage à partir du menu) pour commencer le débogage. Un nouveau workbook Excel s'ouvre avec votre add-in déjà chargé et prêt à l'emploi.

### <a name="start-debugging"></a>Démarrer le débogage

1. Dans VS Code, ouvrez votre fichier de script de code source (**functions.js** ou **functions.ts**).
2. [Définissez un point d'arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.
3. Dans le workbook Excel, entrez une formule qui utilise votre fonction personnalisée.

À ce stade, l'exécution s'arrête sur la ligne de code où vous définissez le point d'arrêt. Vous pouvez désormais vous servir de votre code, définir des montres et utiliser les fonctionnalités de débogage VS Code dont vous avez besoin.

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a>Utiliser le débogger VS Code pour Excel dans Microsoft Edge

Vous pouvez utiliser VS Code pour déboguer des fonctions personnalisées sans interface utilisateur dans Excel dans le navigateur Microsoft Edge. Pour utiliser VS Code avec Microsoft Edge, vous devez installer le [débogger pour l'extension Microsoft Edge.](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge)

### <a name="run-your-add-in-from-vs-code"></a>Exécuter votre add-in à partir de VS Code

1. Ouvrez le dossier de projet racine de vos fonctions personnalisées dans [VS Code.](https://code.visualstudio.com/)
2. Choose **Terminal > Run Task** and type or select **Watch**. Cela surveillera et reconstruira les modifications apportées aux fichiers.
3. Choisissez **Terminal > exécuter la tâche** et tapez ou sélectionnez Serveur **dev.**

### <a name="start-the-vs-code-debugger"></a>Démarrer le débogger VS Code

4. Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.
5. Dans les options Debug, choisissez **Office Online (Edge Chromium).**
6. Ouvrez Excel dans le navigateur Microsoft Edge et créez un nouveau workbook.
7. Choisissez **Partager** dans le ruban et copiez le lien pour l'URL de ce nouveau workbook.
8. Sélectionnez **F5** (ou **exécutez > démarrer** le débogage à partir du menu) pour commencer le débogage. Une invite s'affiche, qui demande l'URL de votre document.
9. Collez l'URL de votre workbook et appuyez sur Entrée.

### <a name="sideload-your-add-in"></a>Charger une version test de votre complément

1. Sélectionnez **l'onglet** Insérer sur le ruban et, dans la section Des **add-ins,** choisissez **Les add-ins Office.**
2. Dans la boîte de dialogue **Des add-ins Office,** sélectionnez l'onglet MES **ADD-INS,** choisissez **Manage My Add-ins**, puis **Upload My Add-in**.
    
    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit indiquant « Gérer mes compléments » et une autre liste déroulante sous cette dernière avec l’option « Charger mon complément »](../images/office-add-ins-my-account.png)

3. **Accédez** au fichier manifeste du add-in, puis sélectionnez **Télécharger.**
    
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)


### <a name="set-breakpoints"></a>Définir des points d'arrêt
1. Dans VS Code, ouvrez votre fichier de script de code source (**functions.js** ou **functions.ts**).
2. [Définissez un point d'arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.
3. Dans le workbook Excel, entrez une formule qui utilise votre fonction personnalisée.

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a>Utiliser les outils de développement de navigateur pour déboguer des fonctions personnalisées dans Excel sur le web

Vous pouvez utiliser les outils de développement du navigateur pour déboguer des fonctions personnalisées sans interface utilisateur dans Excel sur le web. Les étapes suivantes fonctionnent pour Windows et macOS.

### <a name="run-your-add-in-from-visual-studio-code"></a>Exécuter votre add-in à partir de Visual Studio Code

1. Ouvrez le dossier de projet racine de vos fonctions personnalisées [dans Visual Studio Code (VS Code).](https://code.visualstudio.com/)
2. Choose **Terminal > Run Task** and type or select **Watch**. Cela surveillera et reconstruira les modifications apportées aux fichiers.
3. Choose **Terminal > Run Task** and type or select **Dev Server**.

### <a name="sideload-your-add-in"></a>Charger une version test de votre complément

1. Ouvrez [Office sur le web.](https://office.live.com/)
2. Ouvrez un nouveau workbook Excel.
3. Ouvrez **l'onglet** Insérer sur le ruban et, dans la section Des **add-ins,** choisissez **Les add-ins Office.**
4. Dans la boîte de dialogue **Des add-ins Office,** sélectionnez l'onglet MES **ADD-INS,** choisissez **Manage My Add-ins**, puis **Upload My Add-in**.
    
    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit indiquant « Gérer mes compléments » et une autre liste déroulante sous cette dernière avec l’option « Charger mon complément »](../images/office-add-ins-my-account.png)

5. **Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.
    
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

> [!NOTE]
> Une fois que vous avez chargé une version de version sideload dans le document, celui-ci reste chargé de nouveau à chaque ouverture du document.

### <a name="start-debugging"></a>Démarrer le débogage

1. Ouvrez les outils de développement dans le navigateur. Pour Chrome et la plupart des navigateurs F12 ouvrent les outils de développement.
2. Dans les outils de développement, ouvrez votre fichier de script de code source à l'aide de **Cmd+P** ou **Ctrl+P** (**functions.js** ou **functions.ts**).
3. [Définissez un point d'arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée. 

Si vous avez besoin de modifier le code, vous pouvez apporter des modifications dans VS Code et enregistrer les modifications. Actualisez le navigateur pour voir les modifications chargées.

## <a name="use-the-command-line-tools-to-debug"></a>Utiliser les outils de ligne de commande pour déboguer

Si vous n'utilisez pas VS Code, vous pouvez utiliser la ligne de commande (par exemple, Bash ou PowerShell) pour exécuter votre add-in. Vous devez utiliser les outils de développement du navigateur pour déboguer votre code dans Excel sur le web. Vous ne pouvez pas déboguer la version de bureau d'Excel à l'aide de la ligne de commande.

1. À partir de la ligne de commande, `npm run watch` exécutez la commande pour observer et reconstruire lorsque des modifications de code se produisent.
2. Ouvrez une deuxième fenêtre de ligne de commande (la première sera bloquée lors de l'exécution de l'observation).)

3. Si vous souhaitez démarrer votre application dans la version de bureau d'Excel, exécutez la commande suivante :
    
    `npm run start:desktop`
    
    Ou si vous préférez démarrer votre application dans Excel sur le web, exécutez la commande suivante:
    
    `npm run start:web`
    
    Pour Excel sur le web, vous devez également charger une version de version de votre application. Suivez les étapes du chargement de version de version sideload de votre [add-in](#sideload-your-add-in) pour le chargement de version de votre module. Ensuite, continuez jusqu'à la section suivante pour démarrer le débogage.
    
4. Ouvrez les outils de développement dans le navigateur. Pour Chrome et la plupart des navigateurs F12 ouvrent les outils de développement.
5. Dans les outils de développement, ouvrez votre fichier de script de code source (**functions.js** ou **functions.ts**). Votre code de fonctions personnalisées peut se trouver à la fin du fichier.
6. Dans le code source de la fonction personnalisée, appliquez un point d'arrêt en sélectionnant une ligne de code.

Si vous devez modifier le code, vous pouvez effectuer des modifications dans Visual Studio et enregistrer les modifications. Actualisez le navigateur pour voir les modifications chargées.

### <a name="commands-for-building-and-running-your-add-in"></a>Commandes de création et d'exécution de votre add-in

Plusieurs tâches de build sont disponibles :
- `npm run watch`: se construit pour le développement et se reconstruit automatiquement lorsqu'un fichier source est enregistré
- `npm run build-dev`: crée une fois pour le développement
- `npm run build`: builds pour la production
- `npm run dev-server`: exécute le serveur web utilisé pour le développement

Vous pouvez utiliser les tâches suivantes pour démarrer le débogage sur un ordinateur de bureau ou en ligne.
- `npm run start:desktop`: démarre Excel sur le bureau et charge une version de version de votre application.
- `npm run start:web`: démarre Excel sur le web et charge une version de version de votre application.
- `npm run stop`: arrête Excel et le débogage.

## <a name="next-steps"></a>Étapes suivantes
Découvrez les [pratiques d'authentification pour les fonctions personnalisées sans interface utilisateur.](custom-functions-authentication.md)

## <a name="see-also"></a>Voir aussi

* [Résolution des problèmes des fonctions personnalisées](custom-functions-troubleshooting.md)
* [Gestion des erreurs liées aux fonctions personnalisées dans Excel](custom-functions-errors.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
