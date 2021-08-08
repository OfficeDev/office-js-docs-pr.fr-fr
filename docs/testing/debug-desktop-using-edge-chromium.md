---
title: Déboguer des compléments à l’aide de Microsoft Edge WebView2 (avec Chromium)
description: Découvrez comment déboguer un complément Office qui utilise Microsoft Edge WebView2 (avec Chromium) à l’aide du débogueur pour l’extension Microsoft Edge dans VS Code.
ms.date: 07/08/2021
localization_priority: Priority
ms.openlocfilehash: 3a79c9f02aba91d29b3af15a5a3c9113079bc197
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773642"
---
# <a name="debug-add-ins-on-windows-using-edge-chromium-webview2"></a>Déboguer un complément à l’aide de Microsoft Edge WebView2

L’exécution d’un complément Office sur Windows peut utiliser le débogueur pour l’extension Microsoft Edge dans VS Code pour déboguer sur le runtime d’Edge Chromium WebView2.

## <a name="prerequisites"></a>Configuration requise

- [Visual Studio Code](https://code.visualstudio.com/) (doit être exécuté en tant qu’administrateur)
- [Node.js (version 10+)](https://nodejs.org/)
- Windows 10
- [Microsoft Edge Chromium à la disposition des participants au programme Insider de Windows](https://www.microsoftedgeinsider.com/)

## <a name="install-and-use-the-debugger"></a>Installer et utiliser le débogueur

1. Créez un projet à l’aide du [générateur Yoman pour complément Office](https://github.com/OfficeDev/generator-office). Vous pouvez utiliser l’un de nos guides de démarrage rapide, tels que le [Démarrage rapide du complément Outlook](../quickstarts/outlook-quickstart.md) pour pouvoir exécuter cette opération.

> [!TIP]
> Si vous n’utilisez pas de générateur Yeoman basé sur un complément, vous devez régler une clé de Registre. Dans le dossier racine de votre projet, exécutez ce qui suit dans la ligne de commande.
 `office-add-in-debugging start <your manifest path>`

1. Ouvrez le projet dans VS Code. Dans VS Code, sélectionnez **Ctrl + Maj + X** pour ouvrir la barre Extensions. Recherchez l’extension « Débogueur pour Microsoft Edge », puis installez-la.

1. Dans le dossier **.vscode** de votre projet, ouvrez le fichier **launch.json**. Ajoutez le code suivant à la section de configuration.

      ```JSON
        {
          "name": "Debug Office Add-in (Edge Chromium)",
          "type": "edge",
          "request": "attach",
          "useWebView": "advanced",
          "port": 9229,
          "timeout": 600000,
          "webRoot": "${workspaceRoot}",
        },
      ```

1. Ensuite, choisissez **Afficher > Débogage** ou entrez **Ctrl + Maj + D** pour passer à l’affichage Débogage.

1. À partir des options Débogage, choisissez l’option Edge Chromium pour votre application hôte, telle que la **version de bureau d’Excel (Edge Chromium)** Sélectionnez **F5** ou choisissez **Déboguer > Démarrer le débogage** à partir du menu pour commencer le débogage.

1. Dans l’application hôte, telle qu’Excel, votre complément est désormais prêt à être utilisé. Sélectionnez **Afficher le volet de tâches** ou exécutez toute autre commande de complément. Une boîte de dialogue s'affiche, indiquant :

   > Arrêter sur chargement WebView.
   > Pour déboguer l’affichage web, attachez VS Code dans l’instance d’affichage web à l’aide du débogueur Microsoft pour l’extension Edge, puis cliquez sur OK pour continuer. Pour éviter que cette boîte de dialogue ne s’affiche à l’avenir, cliquez sur Annuler.

   Sélectionnez **OK**.

   > [!NOTE]
   > Si vous sélectionnez **Annuler**, la boîte de dialogue ne s’affiche plus lors de l’exécution de cette instance du complément. Toutefois, si vous redémarrez votre complément, la boîte de dialogue s’affichera à nouveau.

1. Vous pourrez définir des points d’arrêt dans le code de votre projet, puis déboguer.

## <a name="see-also"></a>Voir aussi

- [Test et débogage de compléments Office](test-debug-office-add-ins.md)
- [Complément Microsoft Office Extension de débogueur pour Visual Studio Code](debug-with-vs-extension.md)
- [Attacher un débogueur à partir du volet Office](attach-debugger-from-task-pane.md)
