---
title: Complément Microsoft Office Extension de débogueur pour Visual Studio Code
description: Utilisez l’extension Visual Studio code Microsoft Office déboguer votre module de déboguer votre add-in Office.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 60f7e6646cc0bfa2740e3bac0cab5f603b32dd84
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237930"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>Complément Microsoft Office Extension de débogueur pour Visual Studio Code

L’extension Microsoft Office déboguer de l’application pour Visual Studio Code vous permet de déboguer votre application Office par rapport à Microsoft Edge avec le runtime WebView d’origine (EdgeHTML). Pour obtenir des instructions sur le débogage sur Microsoft Edge WebView2 (basé sur Chromium), [consultez cet article.](./debug-desktop-using-edge-chromium.md)

Ce mode de débogage est dynamique, ce qui vous permet de définir des points d’arrêt pendant l’exécution du code. Vous pouvez voir les modifications dans votre code immédiatement lorsque le déboguer est attaché, tout cela sans perdre votre session de débogage. Vos modifications de code sont également persistantes, afin que vous pouvez voir les résultats de plusieurs modifications apportées à votre code. L’image suivante illustre cette extension en action.

![Extension de déboguer du débogage d’une section de modules de débogage de l’extension de débogage de l’extension de débogage d’un addin Office](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>Configuration requise

- [Visual Studio code](https://code.visualstudio.com/) (doit être exécuté en tant qu’administrateur)
- [Node.js (version 10+)](https://nodejs.org/)
- Windows 10
- [Microsoft Edge](https://www.microsoft.com/edge)

Ces instructions supposent que vous avez de l’expérience en utilisant la ligne de commande, que vous comprenez javaScript de base et que vous avez créé un projet de add-in Office avant d’utiliser le générateur Yo Office. Si vous ne l’avez pas encore fait, envisagez de consulter l’un de nos didacticiels, comme ce didacticiel sur les [modules de 2013 excel.](../tutorials/excel-tutorial.md)

## <a name="install-and-use-the-debugger"></a>Installer et utiliser le débogger

1. Si vous devez créer un projet de add-in, [utilisez le générateur Yo Office pour en créer un.](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator) Suivez les invites de la ligne de commande pour configurer votre projet. Vous pouvez choisir n’importe quelle langue ou type de projet en fonction de vos besoins.

> [!NOTE]
> Si vous avez déjà un projet, ignorez l’étape 1 et passez à l’étape 2.

2. Ouvrez une invite de commandes en tant qu’administrateur.
   ![Options d’invite de commandes, y compris « Exécuter en tant qu’administrateur » dans Windows 10](../images/run-as-administrator-vs-code.jpg)

3. Accédez au répertoire de votre projet.

4. Exécutez la commande suivante pour ouvrir votre projet dans Visual Studio Code en tant qu’administrateur.

```command&nbsp;line
code .
```

Une Visual Studio code est ouvert, accédez manuellement au dossier du projet.

> [!TIP]
> Pour ouvrir Visual Studio code en tant qu’administrateur, sélectionnez **l’option** Exécuter en tant qu’administrateur lors de l’ouverture Visual Studio Code après l’avoir recherché dans Windows.

5. Dans VS Code, sélectionnez **Ctrl + Shift + X** pour ouvrir la barre Extensions. Recherchez l’extension « Microsoft Office débompeur de l’extension de module de 2013

6. Dans le dossier .vscode de votre projet, ouvrez le **fichierlaunch.jssur.** Ajoutez le code suivant à la `configurations` section :

```JSON
{
  "type": "office-addin",
  "request": "attach",
  "name": "Attach to Office Add-ins",
  "port": 9222,
  "trace": "verbose",
  "url": "https://localhost:3000/taskpane.html?_host_Info=HOST$Win32$16.01$en-US$$$$0",
  "webRoot": "${workspaceFolder}",
  "timeout": 45000
}
```

7. Dans la section JSON que vous avez copiée, recherchez la section « url ». Dans cette URL, vous devez remplacer le texte HOST en minuscules par l’application qui héberge votre application Office. Par exemple, si votre add-in Office est pour Excel, la valeur de votre URL sera « https://localhost:3000/taskpane.html?_host_Info= <strong>Excel</strong>$Win 32$16.01$en-US$ \$ \$ \$ 0 ».

8. Ouvrez l’invite de commandes et assurez-vous que vous êtes dans le dossier racine de votre projet. Exécutez la commande `npm start` pour démarrer le serveur dev. Lorsque votre add-in se charge dans le client Office, ouvrez le volet Office.

9. Revenir à Visual Studio Code et choisissez Afficher **>** Déboguer ou entrez **Ctrl + Shift + D** pour basculer en mode débogage.

10. Dans les options de débogage, choisissez **Attacher aux add-ins Office.** Sélectionnez **F5** ou **choisissez Déboguer -> démarrer le** débogage à partir du menu pour commencer le débogage.

11. Définissez un point d’arrêt dans le fichier du volet Des tâches de votre projet. Vous pouvez définir des points d’arrêt dans VS Code en pointant sur une ligne de code et en sélectionnant le cercle rouge qui s’affiche.

![Un cercle rouge apparaît sur une ligne de code dans VS Code](../images/set-breakpoint.jpg)

12. Exécutez votre add-in. Vous verrez que les points d’arrêt ont été atteints et que vous pouvez inspecter les variables locales.

## <a name="see-also"></a>Voir aussi

* [Test et débogage de compléments Office](test-debug-office-add-ins.md)

* [Débogage des compléments avec les outils de développement sur Windows 10](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [Déboguer des applications sur Windows à l’aide de Microsoft Edge WebView2 (basé sur Chromium)](debug-desktop-using-edge-chromium.md)
