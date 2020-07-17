---
ms.date: 07/10/2020
description: Découvrez comment déboguer vos fonctions personnalisées Excel qui n’utilisent pas de volet de tâches.
title: Débogage de fonctions personnalisées sans interface utilisateur
localization_priority: Normal
ms.openlocfilehash: 9a493600b6e94d86138cd7949dad0498ec9df05b
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159513"
---
# <a name="ui-less-custom-functions-debugging"></a>Débogage de fonctions personnalisées sans interface utilisateur

Le débogage pour les fonctions personnalisées qui n’utilisent pas de volet de tâches ou d’autres éléments de l’interface utilisateur (fonctions personnalisées sans interface utilisateur) peut être réalisé de plusieurs manières, en fonction de la plateforme que vous utilisez.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

Sur Windows :
- [Débogueur de code Visual Studio et de bureau Excel (code VS)](#use-the-vs-code-debugger-for-excel-desktop)
- [Excel sur le Web et le débogueur de code VS](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [Excel sur le Web et les outils de navigation](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [Ligne de commande](#use-the-command-line-tools-to-debug)

Sur Mac :
- [Excel sur le Web et les outils de navigation](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [Ligne de commande](#use-the-command-line-tools-to-debug)

> [!NOTE]
> Par souci de simplicité, cet article présente le débogage dans le contexte de l’utilisation de Visual Studio code pour modifier, exécuter des tâches et, dans certains cas, utiliser l’affichage débogage. Si vous utilisez un autre éditeur ou outil de ligne de commande, consultez les [instructions de ligne de commande](#commands-for-building-and-running-your-add-in) à la fin de cet article.

## <a name="requirements"></a>Configuration requise

Avant de commencer le débogage, vous devez utiliser le [Générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) afin de créer un projet de fonctions personnalisées. Pour obtenir des instructions sur la création d’un projet de fonctions personnalisées, consultez le didacticiel sur les [fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md).

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>Utiliser le débogueur de code VS pour le bureau Excel

Vous pouvez utiliser le code VS pour déboguer des fonctions personnalisées sans interface utilisateur dans Office Excel sur le bureau.

> [!NOTE]
> Le débogage de bureau pour Mac n’est pas disponible, mais peut être réalisé [à l’aide des outils de navigation et de la ligne de commande pour déboguer Excel sur le Web](#use-the-command-line-tools-to-debug).

### <a name="run-your-add-in-from-vs-code"></a>Exécuter votre complément à partir du code VS

1. Ouvrez votre dossier de projet racine de fonctions personnalisées dans le [code vs](https://code.visualstudio.com/).
2. Choisissez **Terminal > exécuter la tâche** , puis tapez ou sélectionnez **Espion**. Cette opération permet de surveiller et de reconstruire les modifications apportées aux fichiers.
3. Choisissez **Terminal > exécuter la tâche** , puis tapez ou sélectionnez **serveur de développement**.

### <a name="start-the-vs-code-debugger"></a>Démarrer le débogueur de code VS

4. Sélectionnez **afficher > déboguer** ou **Appuyez sur Ctrl + Maj + D** pour basculer vers l’affichage débogage.
5. Dans les options de débogage, choisissez **bureau Excel**.
6. Sélectionnez **F5** (ou choisissez **Déboguer-> démarrer le débogage** dans le menu) pour commencer le débogage. Un nouveau classeur Excel s’ouvre avec votre complément déjà versions test chargées et prêt à être utilisé.

### <a name="start-debugging"></a>Démarrer le débogage

1. Dans le code VS, ouvrez votre fichier de script de code source (**functions.js** ou **functions. TS**).
2. [Définissez un point d’arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.
3. Dans le classeur Excel, entrez une formule qui utilise votre fonction personnalisée.

À ce stade, l’exécution s’arrêtera sur la ligne de code où vous définissez le point d’arrêt. À présent, vous pouvez parcourir votre code, définir des montres et utiliser les fonctionnalités de débogage de code VS dont vous avez besoin.

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a>Utiliser le débogueur de code VS pour Excel dans Microsoft Edge

Vous pouvez utiliser le code VS pour déboguer des fonctions personnalisées sans interface utilisateur dans Excel dans le navigateur Microsoft Edge. Pour utiliser le code VS avec Microsoft Edge, vous devez installer le [débogueur pour l’extension Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) .

### <a name="run-your-add-in-from-vs-code"></a>Exécuter votre complément à partir du code VS

1. Ouvrez votre dossier de projet racine de fonctions personnalisées dans le [code vs](https://code.visualstudio.com/).
2. Choisissez **Terminal > exécuter la tâche** , puis tapez ou sélectionnez **Espion**. Cette opération permet de surveiller et de reconstruire les modifications apportées aux fichiers.
3. Choisissez **Terminal > exécuter la tâche** , puis tapez ou sélectionnez **serveur de développement**.

### <a name="start-the-vs-code-debugger"></a>Démarrer le débogueur de code VS

4. Sélectionnez **afficher > déboguer** ou **Appuyez sur Ctrl + Maj + D** pour basculer vers l’affichage débogage.
5. Dans les options de débogage, sélectionnez **Office Online (Microsoft Edge)**.
6. Ouvrez Excel dans le navigateur Microsoft Edge et créez un classeur.
7. Choisissez **partager** dans le ruban et copiez le lien de l’URL de ce nouveau classeur.
8. Sélectionnez **F5** (ou choisissez **déboguer > démarrer le débogage** dans le menu) pour commencer le débogage. Une invite s’affiche, qui vous demande l’URL de votre document.
9. Collez l’URL de votre classeur, puis appuyez sur entrée.

### <a name="sideload-your-add-in"></a>Charger une version test de votre complément

1. Sélectionnez l’onglet **Insérer** dans le ruban, puis dans la section **compléments** , choisissez **Compléments Office**.
2. Dans la boîte de dialogue **Compléments Office** , sélectionnez l’onglet **mes compléments** , choisissez **gérer mes compléments**, puis **Télécharger mon complément**.
    
    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit indiquant « Gérer mes compléments » et une autre liste déroulante sous cette dernière avec l’option « Charger mon complément »](../images/office-add-ins-my-account.png)

3. **Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.
    
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)


### <a name="set-breakpoints"></a>Définir des points d’arrêt
1. Dans le code VS, ouvrez votre fichier de script de code source (**functions.js** ou **functions. TS**).
2. [Définissez un point d’arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.
3. Dans le classeur Excel, entrez une formule qui utilise votre fonction personnalisée.

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a>Utiliser les outils de développement de navigateur pour déboguer des fonctions personnalisées dans Excel sur le Web

Vous pouvez utiliser les outils de développement de navigateur pour déboguer des fonctions personnalisées sans interface utilisateur dans Excel sur le Web. Les étapes suivantes fonctionnent pour Windows et macOS.

### <a name="run-your-add-in-from-visual-studio-code"></a>Exécuter votre complément à partir de Visual Studio code

1. Ouvrez votre dossier de projet racine de fonctions personnalisées dans [Visual Studio code (vs code)](https://code.visualstudio.com/).
2. Choisissez **Terminal > exécuter la tâche** , puis tapez ou sélectionnez **Espion**. Cette opération permet de surveiller et de reconstruire les modifications apportées aux fichiers.
3. Choisissez **Terminal > exécuter la tâche** , puis tapez ou sélectionnez **serveur de développement**.

### <a name="sideload-your-add-in"></a>Charger une version test de votre complément

1. Ouvrez [Office sur le Web](https://office.live.com/).
2. Ouvrez un nouveau classeur Excel.
3. Ouvrez l’onglet **Insérer** dans le ruban, puis dans la section **compléments** , choisissez **Compléments Office**.
4. Dans la boîte de dialogue **Compléments Office** , sélectionnez l’onglet **mes compléments** , choisissez **gérer mes compléments**, puis **Télécharger mon complément**.
    
    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit indiquant « Gérer mes compléments » et une autre liste déroulante sous cette dernière avec l’option « Charger mon complément »](../images/office-add-ins-my-account.png)

5. **Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.
    
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

> [!NOTE]
> Une fois que vous avez versions test chargées dans le document, il reste versions test chargées chaque fois que vous ouvrez le document.

### <a name="start-debugging"></a>Démarrer le débogage

1. Ouvrez outils de développement dans le navigateur. Pour le chrome et la plupart des navigateurs F12 ouvre les outils de développement.
2. Dans outils de développement, ouvrez votre fichier de script de code source à l’aide de **cmd + p** ou de **Ctrl + p** (**functions.js** ou **functions. TS**).
3. [Définissez un point d’arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée. 

Si vous devez modifier le code, vous pouvez effectuer des modifications dans le code VS et enregistrer les modifications. Actualisez le navigateur pour voir les modifications chargées.

## <a name="use-the-command-line-tools-to-debug"></a>Utiliser les outils de ligne de commande pour déboguer

Si vous n’utilisez pas le code VS, vous pouvez utiliser la ligne de commande (par exemple, bash ou PowerShell) pour exécuter votre complément. Vous devrez utiliser les outils de développement de navigateur pour déboguer votre code dans Excel sur le Web. Vous ne pouvez pas déboguer la version de bureau d’Excel à l’aide de la ligne de commande.

1. À partir de la ligne de commande, exécutez le `npm run watch` suivi et la régénération lorsque les modifications du code se produisent.
2. Ouvrir une deuxième fenêtre de ligne de commande (la première est bloquée lors de l’exécution de la fonction espion).

3. Si vous souhaitez démarrer votre complément dans la version de bureau d’Excel, exécutez la commande suivante :
    
    `npm run start:desktop`
    
    Ou si vous préférez démarrer votre complément dans Excel sur le Web, exécutez la commande suivante :
    
    `npm run start:web`
    
    Pour Excel sur le Web, vous devez également chargement votre complément. Suivez les étapes décrites dans [chargement votre complément](#sideload-your-add-in) pour chargement votre complément. Ensuite, passez à la section suivante pour commencer le débogage.
    
4. Ouvrez outils de développement dans le navigateur. Pour le chrome et la plupart des navigateurs F12 ouvre les outils de développement.
5. Dans outils de développement, ouvrez votre fichier de script de code source (**functions.js** ou **functions. TS**). Votre code de fonctions personnalisées peut être situé à la fin du fichier.
6. Dans le code source de la fonction personnalisée, appliquez un point d’arrêt en sélectionnant une ligne de code.

Si vous devez modifier le code, vous pouvez apporter des modifications dans Visual Studio et enregistrer les modifications. Actualisez le navigateur pour voir les modifications chargées.

### <a name="commands-for-building-and-running-your-add-in"></a>Commandes pour la création et l’exécution de votre complément

Plusieurs tâches de génération sont disponibles :
- `npm run watch`: builds pour le développement et rebuilds automatiques lors de l’enregistrement d’un fichier source
- `npm run build-dev`: builds pour le développement une seule fois
- `npm run build`: builds pour la production
- `npm run dev-server`: exécute le serveur Web utilisé pour le développement

Vous pouvez utiliser les tâches suivantes pour démarrer le débogage sur le bureau ou en ligne.
- `npm run start:desktop`: Démarre Excel sur le bureau et sideloads votre complément.
- `npm run start:web`: Démarre Excel sur le Web et sideloads votre complément.
- `npm run stop`: Arrête Excel et le débogage.

## <a name="next-steps"></a>Étapes suivantes
Découvrez [les pratiques d’authentification pour les fonctions personnalisées sans interface utilisateur](custom-functions-authentication.md).

## <a name="see-also"></a>Consultez également

* [Dépannage des fonctions personnalisées](custom-functions-troubleshooting.md)
* [Gestion des erreurs liées aux fonctions personnalisées dans Excel](custom-functions-errors.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
