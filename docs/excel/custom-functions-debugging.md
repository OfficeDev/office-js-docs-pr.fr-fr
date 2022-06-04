---
title: Débogage de fonctions personnalisées sans interface utilisateur
description: Découvrez comment déboguer vos fonctions personnalisées Excel qui n’utilisent pas de volet Office.
ms.date: 05/31/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1268aa07d160970fda12f8fccbe88e0427b246fc
ms.sourcegitcommit: 81f6018ac9731ff73e36d30f5ff10df21504c093
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/04/2022
ms.locfileid: "65891927"
---
# <a name="ui-less-custom-functions-debugging"></a>Débogage de fonctions personnalisées sans interface utilisateur

Cet article traite du débogage *uniquement* pour les fonctions personnalisées qui n’utilisent pas de volet Office ou d’autres éléments d’interface utilisateur (fonctions personnalisées sans interface utilisateur).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

Sur Windows :

- [Débogueur Excel Desktop et Visual Studio Code (VS Code)](#use-the-vs-code-debugger-for-excel-desktop)
- [Débogueur Excel sur le web et VS Code](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [Excel sur le web et les outils de navigateur](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [Ligne de commande](#use-the-command-line-tools-to-debug)

Sur Mac :

- [Excel sur le web et les outils de navigateur](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [Ligne de commande](#use-the-command-line-tools-to-debug)

> [!NOTE]
> Par souci de simplicité, cet article présente le débogage dans le contexte de l’utilisation de Visual Studio Code pour modifier, exécuter des tâches et, dans certains cas, utiliser la vue de débogage. Si vous utilisez un autre éditeur ou outil en ligne de commande, consultez les [instructions de ligne de commande](#commands-for-building-and-running-your-add-in) à la fin de cet article.

## <a name="requirements"></a>Conditions requises

Ce processus de débogage fonctionne **uniquement** pour les fonctions personnalisées sans interface utilisateur, qui n’utilisent pas de volet Office ou d’autres éléments d’interface utilisateur. Pour créer une fonction personnalisée sans interface utilisateur, suivez les étapes du didacticiel [Créer des fonctions personnalisées dans Excel](../tutorials/excel-tutorial-create-custom-functions.md) , puis supprimez tous les éléments du volet Office et de l’interface utilisateur installés par le [générateur Yeoman pour compléments Office](../develop/yeoman-generator-overview.md).

Notez que ce processus de débogage n’est pas compatible avec les projets de fonctions personnalisées utilisant un [runtime partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

Ce processus de débogage ne fonctionne pas avec les projets créés avec le **projet de complément Office contenant l’option manifeste uniquement** dans le générateur Yeoman. Les scripts mentionnés plus loin dans cet article ne sont pas installés avec cette option. Pour déboguer un complément créé avec cette option, consultez les instructions de l’un de ces articles, le cas échéant. 

- [Déboguer des compléments à l’aide des Outils de développement dans Microsoft Edge (basés sur Chromium)](../testing/debug-add-ins-using-devtools-edge-chromium.md)
- [Déboguer des compléments à l’aide d’outils de développement dans Internet Explorer](../testing/debug-add-ins-using-f12-tools-ie.md)
- [Déboguer des compléments Office sur un Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>Utiliser le débogueur VS Code pour Excel Desktop

Vous pouvez utiliser VS Code pour déboguer des fonctions personnalisées sans interface utilisateur dans Office Excel sur le bureau.

> [!NOTE]
> Le débogage de bureau pour Mac n’est pas disponible, mais peut être réalisé [à l’aide des outils de navigateur et de la ligne de commande pour déboguer Excel sur le web](#use-the-command-line-tools-to-debug)).

### <a name="run-your-add-in-from-vs-code"></a>Exécuter votre complément à partir de VS Code

1. Ouvrez le dossier de projet racine de vos fonctions personnalisées dans [VS Code](https://code.visualstudio.com/).
1. Choisissez **Terminal > Exécuter la tâche** et tapez ou sélectionnez **Espion**. Cela permet de surveiller et de reconstruire les modifications apportées aux fichiers.
1. Choisissez **Terminal > Exécuter la tâche** et tapez ou sélectionnez **Serveur de développement**.

### <a name="start-the-vs-code-debugger"></a>Démarrer le débogueur VS Code

1. Choisissez **Afficher > Exécuter** ou entrez **Ctrl+Maj+D** pour basculer en mode débogage.
1. Dans le menu déroulant Exécuter, choisissez **Excel Desktop (Fonctions personnalisées).**
1. Sélectionnez **F5** (ou **sélectionnez Exécuter -> Démarrer le débogage** dans le menu) pour commencer le débogage. Un nouveau classeur Excel s’ouvre avec votre complément déjà chargé de manière indépendante et prêt à l’emploi.

### <a name="start-debugging"></a>Démarrer le débogage

1. Dans VS Code, ouvrez votre fichier de script de code source (**functions.js** ou **functions.ts**).
2. [Définissez un point d’arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.
3. Dans le classeur Excel, entrez une formule qui utilise votre fonction personnalisée.

À ce stade, l’exécution s’arrête sur la ligne de code où vous définissez le point d’arrêt. Vous pouvez maintenant parcourir votre code, définir des montres et utiliser toutes les fonctionnalités de débogage VS Code dont vous avez besoin.

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a>Utiliser le débogueur VS Code pour Excel dans Microsoft Edge

Vous pouvez utiliser VS Code pour déboguer des fonctions personnalisées sans interface utilisateur dans Excel sur le navigateur Microsoft Edge. Pour utiliser VS Code avec Microsoft Edge, vous devez installer [l’extension Microsoft Edge DevTools pour Visual Studio Code](/microsoft-edge/visual-studio-code/microsoft-edge-devtools-extension).

### <a name="run-your-add-in-from-vs-code"></a>Exécuter votre complément à partir de VS Code

1. Ouvrez le dossier de projet racine de vos fonctions personnalisées dans [VS Code](https://code.visualstudio.com/).
2. Choisissez **Terminal > Exécuter la tâche** et tapez ou sélectionnez **Espion**. Cela permet de surveiller et de reconstruire les modifications apportées aux fichiers.
3. Choisissez **Terminal > Exécuter la tâche** et tapez ou sélectionnez **Serveur de développement**.

### <a name="start-the-vs-code-debugger"></a>Démarrer le débogueur VS Code

1. Choisissez **Afficher > Exécuter** ou entrez **Ctrl+Maj+D** pour basculer en mode débogage.
1. Dans les options de débogage, choisissez **Office Online (Edge Chromium).**
1. Ouvrez Excel dans le navigateur Microsoft Edge et créez un classeur.
1. Choisissez **Partager** dans le ruban et copiez le lien pour l’URL de ce nouveau classeur.
1. Sélectionnez **F5** (ou **sélectionnez Exécuter > Démarrer le débogage** dans le menu) pour commencer le débogage. Une invite s’affiche, qui demande l’URL de votre document.
1. Collez l’URL de votre classeur, puis appuyez sur Entrée.

### <a name="sideload-your-add-in"></a>Charger une version test de votre complément

1. Sélectionnez l’onglet **Insertion** dans le ruban et, dans la section **Compléments** , choisissez **Compléments Office**.
2. Dans la boîte **de dialogue Compléments Office** , sélectionnez l’onglet **MES COMPLÉMENTS** , choisissez **Gérer mes compléments**, puis **chargez mon complément**.
  
    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit en lisant « Gérer mes compléments » et une liste déroulante en dessous avec l’option « Télécharger mon complément ».](../images/office-add-ins-my-account.png)

3. **Accédez** au fichier manifeste du complément, puis **sélectionnez Charger**.
  
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

### <a name="set-breakpoints"></a>Définir des points d’arrêt

1. Dans VS Code, ouvrez votre fichier de script de code source (**functions.js** ou **functions.ts**).
2. [Définissez un point d’arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.
3. Dans le classeur Excel, entrez une formule qui utilise votre fonction personnalisée.

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a>Utiliser les outils de développement du navigateur pour déboguer des fonctions personnalisées dans Excel sur le web

Vous pouvez utiliser les outils de développement du navigateur pour déboguer des fonctions personnalisées sans interface utilisateur dans Excel sur le web. Les étapes suivantes fonctionnent pour Windows et macOS.

### <a name="run-your-add-in-from-visual-studio-code"></a>Exécuter votre complément à partir de Visual Studio Code

1. Ouvrez le dossier de projet racine de vos fonctions personnalisées dans [Visual Studio Code (VS Code).](https://code.visualstudio.com/)
2. Choisissez **Terminal > Exécuter la tâche** et tapez ou sélectionnez **Espion**. Cela permet de surveiller et de reconstruire les modifications apportées aux fichiers.
3. Choisissez **Terminal > Exécuter la tâche** et tapez ou sélectionnez **Serveur de développement**.

### <a name="sideload-your-add-in"></a>Charger une version test de votre complément

1. Ouvrez [Office sur le web](https://office.live.com/).
2. Ouvrez un nouveau classeur Excel.
3. Ouvrez l’onglet **Insertion** dans le ruban et, dans la section **Compléments** , choisissez **Compléments Office**.
4. Dans la boîte **de dialogue Compléments Office** , sélectionnez l’onglet **MES COMPLÉMENTS** , choisissez **Gérer mes compléments**, puis **chargez mon complément**.
  
    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit en lisant « Gérer mes compléments » et une liste déroulante en dessous avec l’option « Télécharger mon complément ».](../images/office-add-ins-my-account.png)

5. **Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.
  
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

> [!NOTE]
> Une fois que vous avez chargé une version test du document, celui-ci reste chargé de manière indépendante chaque fois que vous ouvrez le document.

### <a name="start-debugging"></a>Démarrer le débogage

1. Ouvrez les outils de développement dans le navigateur. Pour Chrome et la plupart des navigateurs, F12 ouvre les outils de développement.
2. Dans les outils de développement, ouvrez votre fichier de script de code source à l’aide de **Cmd+P** ou **Ctrl+P** (**functions.js** ou **functions.ts**).
3. [Définissez un point d’arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée. 

Si vous devez modifier le code, vous pouvez apporter des modifications dans VS Code et enregistrer les modifications. Actualisez le navigateur pour voir les modifications chargées.

## <a name="use-the-command-line-tools-to-debug"></a>Utiliser les outils en ligne de commande pour déboguer

Si vous n’utilisez pas VS Code, vous pouvez utiliser la ligne de commande (par exemple Bash ou PowerShell) pour exécuter votre complément. Vous devez utiliser les outils de développement du navigateur pour déboguer votre code dans Excel sur le web. Vous ne pouvez pas déboguer la version de bureau d’Excel à l’aide de la ligne de commande.

1. À partir de la ligne de commande, exécutez `npm run watch` pour surveiller et régénérer quand des modifications de code se produisent.
2. Ouvrez une deuxième fenêtre de ligne de commande (la première sera bloquée lors de l’exécution de la montre.)

3. Si vous souhaitez démarrer votre complément dans la version de bureau d’Excel, exécutez la commande suivante.
  
    `npm run start:desktop`
  
    Ou si vous préférez démarrer votre complément dans Excel sur le web, exécutez la commande suivante.
  
    `npm run start:web -- --document {url}` (où `{url}` se trouve l’URL d’un fichier Excel sur OneDrive ou SharePoint)
  
    Si votre complément ne charge pas de version test dans le document, suivez les étapes décrites dans Chargement indépendant de [votre complément](#sideload-your-add-in) pour charger de manière test votre complément. Passez ensuite à la section suivante pour commencer le débogage.
  
4. Ouvrez les outils de développement dans le navigateur. Pour Chrome et la plupart des navigateurs, F12 ouvre les outils de développement.
5. Dans les outils de développement, ouvrez votre fichier de script de code source (**functions.js** ou **functions.ts**). Le code de vos fonctions personnalisées peut se trouver à la fin du fichier.
6. Dans le code source de la fonction personnalisée, appliquez un point d’arrêt en sélectionnant une ligne de code.

Si vous devez modifier le code, vous pouvez apporter des modifications dans Visual Studio et enregistrer les modifications. Actualisez le navigateur pour voir les modifications chargées.

### <a name="commands-for-building-and-running-your-add-in"></a>Commandes pour la création et l’exécution de votre complément

Plusieurs tâches de génération sont disponibles.

- `npm run watch`: génère pour le développement et se reconstruit automatiquement lorsqu’un fichier source est enregistré
- `npm run build-dev`: builds pour le développement une seule fois
- `npm run build`: builds pour la production
- `npm run dev-server`: exécute le serveur web utilisé pour le développement

Vous pouvez utiliser les tâches suivantes pour démarrer le débogage sur le bureau ou en ligne.

- `npm run start:desktop`: démarre Excel sur le bureau et charge de manière indépendante votre complément.
- `npm run start:web -- --document {url}` (où `{url}` se trouve l’URL d’un fichier Excel sur OneDrive ou SharePoint) : démarre Excel sur le web et charge de manière indépendante votre complément.
- `npm run stop`: arrête Excel et le débogage.

## <a name="next-steps"></a>Étapes suivantes

Découvrez les [pratiques d’authentification pour les fonctions personnalisées sans interface utilisateur](custom-functions-authentication.md).

## <a name="see-also"></a>Voir aussi

* [Résolution des problèmes liés aux fonctions personnalisées](custom-functions-troubleshooting.md)
* [Gestion des erreurs liées aux fonctions personnalisées dans Excel](custom-functions-errors.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
