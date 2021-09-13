---
title: Déboguer des compléments à l’aide de Microsoft Edge WebView2 (avec Chromium)
description: Découvrez comment déboguer un complément Office qui utilise Microsoft Edge WebView2 (avec Chromium) à l’aide du débogueur pour l’extension Microsoft Edge dans VS Code.
ms.date: 08/18/2021
ms.localizationpriority: high
ms.openlocfilehash: 90c58352ef49a365cd80dd7699d23f40c800bbd6
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149074"
---
# <a name="debug-add-ins-on-windows-using-edge-chromium-webview2"></a>Déboguer un complément à l’aide de Microsoft Edge WebView2

L’exécution d’un complément Office sur Windows peut utiliser le débogueur pour l’extension Microsoft Edge dans VS Code pour déboguer sur le runtime d’Edge Chromium WebView2.

## <a name="prerequisites"></a>Configuration requise

- [Visual Studio Code](https://code.visualstudio.com/) (doit être exécuté en tant qu’administrateur)
- [Node.js (version 10+)](https://nodejs.org/)
- Windows 10
- La combinaison d’une plateforme et d’une application Office qui prend en charge Microsoft Edge avec WebView2 (basé sur Chromium), comme expliqué dans [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md). Si votre version de Microsoft 365 est antérieure à 2101, vous devez installer WebView2. Suivez les instructions pour l’installer sur [Microsoft Edge WebView2/Incorporer du contenu web... avec Microsoft Edge webView2](https://developer.microsoft.com/microsoft-edge/webview2/).

## <a name="install-and-use-the-debugger"></a>Installer et utiliser le débogueur

1. Créez un projet à l’aide du [générateur Yoman pour complément Office](https://github.com/OfficeDev/generator-office). Vous pouvez utiliser l’un de nos guides de démarrage rapide, tels que le [Démarrage rapide du complément Outlook](../quickstarts/outlook-quickstart.md) pour pouvoir exécuter cette opération.

    > [!TIP]
    > Si vous n’utilisez pas de générateur Yeoman basé sur un complément, vous pouvez être invité à ajuster une clé de Registre. Dans le dossier racine de votre projet, exécutez ce qui suit en ligne de commande.`office-add-in-debugging start <your manifest path>`

1. Ouvrez le projet dans VS Code. Dans VS Code, sélectionnez **Ctrl + Shift + X** pour ouvrir la barre Extensions. Recherchez l’extension « Débogueur pour Microsoft Edge », puis installez-la.

1. Ensuite, choisissez  **Afficher > Exécuter** ou entrez **Ctrl+Shift+D** pour basculer en mode débogage.

1. Dans les options **EXECUTER ET DEBOGUER**, choisissez l’option Edge Chromium pour votre application hôte, telle que **Version de bureau d’Excel (Edge Chromium)**. Sélectionnez **F5** ou choisissez **Exécuter > Démarrer le débogage** dans le menu pour commencer le débogage. Cette action lance automatiquement un serveur local dans une fenêtre Node pour héberger votre complément, puis ouvre automatiquement l’application hôte, telle qu’Excel ou Word. Cela peut prendre plusieurs heures.

1. Dans l’application hôte, votre complément est désormais prêt à être utilisé. Sélectionnez **Afficher le volet de tâches** ou exécutez toute autre commande de complément. Une boîte de dialogue s'affiche, indiquant :

   > Arrêter sur chargement WebView.
   > Pour déboguer l’affichage web, attachez VS Code dans l’instance d’affichage web à l’aide du débogueur Microsoft pour l’extension Edge, puis cliquez sur OK pour continuer. Pour éviter que cette boîte de dialogue ne s’affiche à l’avenir, cliquez sur Annuler.

   Sélectionnez **OK**.

   > [!NOTE]
   > Si vous sélectionnez **Annuler**, la boîte de dialogue ne s’affiche plus lors de l’exécution de cette instance du complément. Toutefois, si vous redémarrez votre complément, la boîte de dialogue s’affichera à nouveau.

1. Vous pourrez définir des points d’arrêt dans le code de votre projet, puis déboguer.

   > [!NOTE]
   > Les points d’arrêt dans les appels de `Office.initialize` ou de `Office.onReady` sont ignorés. Pour plus d’informations sur ces méthodes, consultez [Initialiser votre complément Office](../develop/initialize-add-in.md).

> [!IMPORTANT]
> La meilleure façon d’arrêter une session de débogage consiste à sélectionner **Shift+F5** ou à choisir **Exécuter > Arrêter le débogage** dans le menu. Cette action doit fermer la fenêtre du serveur Node et tenter de fermer l’application hôte, mais une invite s’affiche sur l’application hôte vous demandant s’il faut enregistrer le document ou non. Faites un choix approprié et laissez l’application hôte se fermer. Évitez de fermer manuellement la fenêtre Node ou l’application hôte. Cela peut entraîner des bogues en particulier lorsque vous arrêtez et démarrez des sessions de débogage à plusieurs reprises.
>
> Si le débogage cesse de fonctionner ; par exemple, si les points d’arrêt sont ignorés ; arrêter le débogage. Ensuite, si nécessaire, fermez toutes les fenêtres d’application hôte et la fenêtre Nœud. Enfin, fermez Visual Studio Code et rouvrez-le.

## <a name="see-also"></a>Voir aussi

- [Test et débogage de compléments Office](test-debug-office-add-ins.md)
- [Complément Microsoft Office Extension de débogueur pour Visual Studio Code](debug-with-vs-extension.md)
- [Attacher un débogueur à partir du volet Office](attach-debugger-from-task-pane.md)
