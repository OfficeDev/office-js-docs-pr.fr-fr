---
title: Déboguer les add-ins sur Windows à l'aide de Visual Studio Code et du WebView hérité de Microsoft Edge (EdgeHTML)
description: Découvrez comment déboguer des Office qui utilisent Version antérieure de Microsoft Edge WebView (EdgeHTML) à l’aide de l’extension de déboguer du Office dans VS Code.
ms.date: 02/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 98ac2fc4eab8ef20ddc416b299fde88095ef4fb4
ms.sourcegitcommit: 4a7b9b9b359d51688752851bf3b41b36f95eea00
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/22/2022
ms.locfileid: "63710943"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>Complément Microsoft Office Extension de débogueur pour Visual Studio Code

Office les applications qui s’exécutent sur Windows peuvent utiliser l’extension déboguer le déboguer de Office dans Visual Studio Code pour déboguer sur Version antérieure de Microsoft Edge avec le runtime WebView d’origine (EdgeHTML). 

> [!IMPORTANT]
> Cet article s’applique uniquement lorsque Office exécute des applications dans le runtime WebView (EdgeHTML) d’origine, comme expliqué dans les [navigateurs](../concepts/browsers-used-by-office-web-add-ins.md) utilisés par les Office. Pour obtenir des instructions sur le débogage dans le code Visual Studio par rapport à Microsoft Edge WebView2 (basé sur Chromium), voir Microsoft Office [Add-in Debugger Extension for Visual Studio Code](debug-desktop-using-edge-chromium.md).

> [!TIP]
> Si vous ne pouvez pas ou ne souhaitez pas déboguer à l’aide d’outils intégrés à Visual Studio Code ; ou si vous rencontrez un problème qui se produit uniquement lorsque le module est exécuté en dehors de Visual Studio Code, vous pouvez déboguer le runtime Edge Legacy (EdgeHTML) à l’aide des outils de développement edge hérités, comme décrit dans Débogage des applications à l’aide des [outils de développement dans Version antérieure de Microsoft Edge](debug-add-ins-using-devtools-edge-legacy.md).

Ce mode de débogage est dynamique et vous permet de définir des points d'arrêt pendant l'exécution du code. Vous pouvez voir les modifications dans votre code immédiatement lorsque le déboguer est attaché, tout cela sans perdre votre session de débogage. Vos modifications de code sont également persistantes, afin que vous pouvez voir les résultats de plusieurs modifications apportées à votre code. L’image suivante illustre cette extension en action.

![Office extension déboguer une section de Excel de débogage.](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>Conditions préalables

- [Visual Studio Code](https://code.visualstudio.com/)
- [Node.js (version 10+)](https://nodejs.org/)
- Windows 10, 11
- [Microsoft Edge](https://www.microsoft.com/edge) Combinaison de plateforme et d’application Office qui prend en charge les Version antérieure de Microsoft Edge avec le webview d’origine (EdgeHTML), comme expliqué dans les [navigateurs](../concepts/browsers-used-by-office-web-add-ins.md) utilisés par les Office de contenu.

## <a name="install-and-use-the-debugger"></a>Installer et utiliser le débogueur

Ces instructions supposent que vous avez de l’expérience en utilisant la ligne de commande, que vous comprenez javaScript de base et que vous avez créé un projet de Office Add-in avant d’utiliser le générateur [Yeoman](../develop/yeoman-generator-overview.md) pour les Office. Si vous ne l’avez pas encore fait, envisagez de consulter l’un de nos didacticiels, comme Excel Office [didacticiel sur le Excel Office de l’aide](../tutorials/excel-tutorial.md).

1. La première étape dépend du projet et de la façon dont il a été créé.

   - Si vous souhaitez créer un projet pour expérimenter le débogage dans Visual Studio Code, utilisez [le générateur Yeoman pour les modules complémentaires Office](../develop/yeoman-generator-overview.md). Pour ce faire, utilisez l'un de nos guides de démarrage rapide, tel que [le guide de démarrage rapide du module complémentaire Outlook](../quickstarts/outlook-quickstart.md). 
   - Si vous souhaitez déboguer un projet existant créé avec Yo Office, passez à l’étape suivante.
   - Si vous souhaitez déboguer un projet existant qui n’a pas été créé avec Yo Office, effectuez la procédure dans l’annexe, puis revenir à l’étape suivante de cette procédure.[](#appendix)


1. Ouvrez VS Code et ouvrez votre projet dans celui-ci. 

1. Dans VS Code, sélectionnez **Ctrl + Shift + X** pour ouvrir la barre Extensions. Recherchez l’extension « Microsoft Office débompeur de l’extension de la mise en Microsoft Office » et installez-la.

1. Choisissez **View > Run** ou entrez **Ctrl+Shift+D** pour passer en mode débogage.

1. Dans les options **EXÉCUTER ET DÉBOGUER**, choisissez l’option Edge héritée pour votre application hôte, par exemple Outlook **Desktop (Edge Legacy).** Sélectionnez **F5** ou choisissez **Exécuter > Démarrer le débogage** dans le menu pour commencer le débogage. Cette action lance automatiquement un serveur local dans une fenêtre Node pour héberger votre complément, puis ouvre automatiquement l’application hôte, telle qu’Excel ou Word. Cela peut prendre plusieurs heures.

1. Dans l’application hôte, votre complément est désormais prêt à être utilisé. Sélectionnez **Afficher le volet de tâches** ou exécutez toute autre commande de complément. Une boîte de dialogue s’affiche comme suit :

   > Arrêter sur chargement WebView.
   > Pour déboguer le WebView, attachez VS Code à l’instance WebView à l’aide du déboguer Microsoft pour l’extension Edge, puis cliquez sur **OK** pour continuer. Pour empêcher cette boîte de dialogue d’apparaître à l’avenir, cliquez sur **Annuler**.

   Sélectionnez **OK**.

   > [!NOTE]
   > Si vous sélectionnez **Annuler**, la boîte de dialogue ne s’affiche plus lors de l’exécution de cette instance du complément. Toutefois, si vous redémarrez votre complément, la boîte de dialogue s’affichera à nouveau.

1. Définissez un point d’arrêt dans le fichier du volet Des tâches de votre projet. Pour définir des points d'arrêt dans Visual Studio Code, passez la souris à côté d'une ligne de code et sélectionnez le cercle rouge qui apparaît.

    ![Un cercle rouge apparaît sur une ligne de code Visual Studio Code.](../images/set-breakpoint.jpg)

1. Exécutez la fonctionnalité dans votre add-in qui appelle les lignes avec des points d'arrêt. Vous verrez que les points d’arrêt ont été atteints et que vous pouvez inspecter les variables locales.

   > [!NOTE]
   > Les points d’arrêt dans les appels de `Office.initialize` ou de `Office.onReady` sont ignorés. Pour plus d’informations sur ces méthodes, consultez [Initialiser votre complément Office](../develop/initialize-add-in.md).

> [!IMPORTANT]
> La meilleure façon d’arrêter une session de débogage consiste à sélectionner **Shift+F5** ou à choisir **Exécuter > Arrêter le débogage** dans le menu. Cette action doit fermer la fenêtre du serveur Node et tenter de fermer l’application hôte, mais une invite s’affiche sur l’application hôte vous demandant s’il faut enregistrer le document ou non. Faites un choix approprié et laissez l’application hôte se fermer. Évitez de fermer manuellement la fenêtre Node ou l’application hôte. Cela peut entraîner des bogues en particulier lorsque vous arrêtez et démarrez des sessions de débogage à plusieurs reprises.
>
> Si le débogage cesse de fonctionner ; par exemple, si les points d’arrêt sont ignorés ; arrêter le débogage. Ensuite, si nécessaire, fermez toutes les fenêtres d’application hôte et la fenêtre Nœud. Enfin, fermez Visual Studio Code et rouvrez-le.

### <a name="appendix"></a>Annexe

Si votre projet n’a pas été créé avec Yo Office, vous devez créer une configuration de débogage pour Visual Studio Code. 

1. Créez un fichier nommé `launch.json` dans le dossier du projet `\.vscode` s'il n'y en a pas déjà un. 
1. Assurez-vous que le fichier possède un `configurations` tableau. Voici un exemple simple d’un `launch.json`.

    ```json
    {
      // other properities may be here.

      "configurations": [

        // configuration objects may be here.

      ]

      //other properies may be here.
    }
    ```

1. Ajoutez l’objet suivant au `configurations` tableau.

    ```json
    {
      "name": "HOST Desktop (Edge Legacy)",
      "type": "office-addin",
      "request": "attach",
      "url": "https://localhost:3000/taskpane.html?_host_Info=HOST$Win32$16.01$en-US$$$$0",
      "port": 9222,
      "timeout": 600000,
      "webRoot": "${workspaceRoot}",
      "preLaunchTask": "Debug: HOST Desktop",
      "postDebugTask": "Stop Debug"
    }
    ```

1. Remplacez l’espace `HOST` réservé aux trois endroits par le nom de l’application Office dans qui s’exécute le add-in ; par exemple, `Outlook` ou `Word`.
1. Enregistrez et fermez le fichier.

## <a name="see-also"></a>Voir aussi

- [Test et débogage de compléments Office](test-debug-office-add-ins.md)
- [Déboguer des Windows à l’Visual Studio Code et Microsoft Edge WebView2 (basé sur Chromium)](debug-desktop-using-edge-chromium.md).
- [Déboguer des compléments à l’aide des outils de développement pour Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
- [Déboguer des compléments à l’aide des outils de développement pour la version héritée Edge](debug-add-ins-using-devtools-edge-legacy.md)
- [Déboguer des compléments à l’aide des Outils de développement dans Microsoft Edge (basés sur Chromium)](debug-add-ins-using-devtools-edge-chromium.md)
- [Attacher un débogueur à partir du volet Office](attach-debugger-from-task-pane.md)
