---
title: Complément Microsoft Office Extension de débogueur pour Visual Studio Code
description: Utilisez le débogueur de complément Microsoft Office de l’extension de code Visual Studio pour déboguer votre complément Office.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 2439af12f30cef1b9d291578cbababe3ed601644
ms.sourcegitcommit: 7d5407d3900d2ad1feae79a4bc038afe50568be0
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/30/2020
ms.locfileid: "46530470"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>Complément Microsoft Office Extension de débogueur pour Visual Studio Code

L’extension du débogueur de complément Microsoft Office pour Visual Studio code vous permet de déboguer votre complément Office par rapport au runtime Edge.

Ce mode de débogage est dynamique, ce qui vous permet de définir des points d’arrêt lors de l’exécution du code. Vous pouvez voir les modifications apportées à votre code immédiatement lorsque le débogueur est attaché, tout cela sans perdre votre session de débogage. Les modifications apportées au code sont également conservées, ce qui vous permet de voir les résultats de plusieurs modifications apportées à votre code. L’image suivante illustre cette extension en action.

![Extension de débogage du complément Office AddIn débogage d’une section de compléments Excel](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>Conditions préalables

- [Visual Studio code](https://code.visualstudio.com/) (doit être exécuté en tant qu’administrateur)
- [Node.js (version 10 +)](https://nodejs.org/)
- Windows 10
- [Microsoft Edge](https://www.microsoft.com/edge)

Ces instructions supposent que vous avez une expérience en utilisant la ligne de commande, que vous compreniez JavaScript de base et que vous avez créé un projet de complément Office avant d’utiliser le générateur Yo Office. Si vous ne l’avez pas encore fait, songez à consulter l’un de nos didacticiels, comme le [didacticiel sur les compléments Office Excel](../tutorials/excel-tutorial.md).

## <a name="install-and-use-the-debugger"></a>Installer et utiliser le débogueur

1. Si vous devez créer un projet de complément, [Utilisez le générateur Yo Office pour en créer un](https://docs.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator). Suivez les invites de la ligne de commande pour configurer votre projet. Vous pouvez choisir n’importe quelle langue ou type de projet en fonction de vos besoins.

> [!NOTE]
> Si vous disposez déjà d’un projet, ignorez l’étape 1 et passez à l’étape 2.

2. Ouvrez une invite de commandes en tant qu’administrateur.
   ![Options d’invite de commandes, y compris « exécuter en tant qu’administrateur » dans Windows 10](../images/run-as-administrator-vs-code.jpg)

3. Naviguez jusqu’au répertoire de votre projet.

4. Exécutez la commande suivante pour ouvrir votre projet dans Visual Studio code en tant qu’administrateur.

```command&nbsp;line
code .
```

Une fois Visual Studio code ouvert, accédez manuellement au dossier du projet.

> [!TIP]
> Pour ouvrir Visual Studio code en tant qu’administrateur, sélectionnez l’option **exécuter en tant qu’administrateur** lors de l’ouverture de Visual Studio code après avoir effectué une recherche dans Windows.

5. Dans le code VS, sélectionnez **Ctrl + Maj + X** pour ouvrir la barre extensions. Recherchez l’extension « Microsoft Office Add-in Debugger » et installez-la.

6. Dans le dossier. vscode de votre projet, ouvrez le fichier **launch.js** . Ajoutez le code suivant à la `configurations` section :

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

7. Dans la section de JSON que vous venez de copier, recherchez la section « URL ». Dans cette URL, vous devrez remplacer le texte d’hôte en majuscules par l’application hôte pour votre complément Office. Par exemple, si votre complément Office est destiné à Excel, la valeur de votre URL serait « https://localhost:3000/taskpane.html?_host_Info= <strong>Excel</strong>$Win 32 $16.01 $ en-US $ \$ \$ \$ 0 ».

8. Ouvrez l’invite de commandes et assurez-vous que vous vous trouvez dans le dossier racine de votre projet. Exécutez la commande `npm start` pour démarrer le serveur de développement. Lorsque votre complément est chargé dans le client Office, ouvrez le volet de tâches.

9. Revenez à Visual Studio code et choisissez **view > Debug** ou **Appuyez sur Ctrl + Maj + D** pour basculer vers le mode débogage.

10. Dans les options de débogage, choisissez **attacher aux compléments Office**. Sélectionnez **F5** ou choisissez **Déboguer-> démarrer le débogage** dans le menu pour commencer le débogage.

11. Définissez un point d’arrêt dans le fichier de volet Office de votre projet. Vous pouvez définir des points d’arrêt dans le code VS en plaçant le curseur en regard d’une ligne de code et en sélectionnant le cercle rouge qui apparaît.

![Un cercle rouge apparaît sur une ligne de code dans un code VS](../images/set-breakpoint.jpg)

12. Exécutez votre complément. Vous verrez que des points d’arrêt ont été atteints et que vous pouvez inspecter les variables locales.

## <a name="see-also"></a>Voir aussi

* [Test et débogage de compléments Office](test-debug-office-add-ins.md)

* [Débogage des compléments avec les outils de développement sur Windows 10](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [Attacher un débogueur à partir du volet Office](attach-debugger-from-task-pane.md)
