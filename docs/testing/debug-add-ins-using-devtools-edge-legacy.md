---
title: Déboguer des applications à l’aide des outils de développement pour Version antérieure de Microsoft Edge
description: Déboguer des applications à l’aide des outils de développement Version antérieure de Microsoft Edge.
ms.date: 11/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: e3d0b77a6898dcefc7fba7c9d52eb739a2d685aa
ms.sourcegitcommit: a3debae780126e03a1b566efdec4d8be83e405b8
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/03/2021
ms.locfileid: "60809076"
---
# <a name="debug-add-ins-using-developer-tools-in-microsoft-edge-legacy"></a>Déboguer des applications à l’aide des outils de développement Version antérieure de Microsoft Edge

Cet article montre comment déboguer le code côté client (JavaScript ou TypeScript) de votre add-in lorsque les conditions suivantes sont remplies.

- Vous ne pouvez pas ou ne souhaitez pas déboguer à l’aide des outils intégrés à votre IDE ; ou vous rencontrez un problème qui se produit uniquement lorsque le module est exécuté en dehors de l’IDE.
- Votre ordinateur utilise une combinaison de versions Windows et Office qui utilisent le contrôle webview Edge d’origine, EdgeHTML.

> [!TIP]
> Pour plus d’informations sur le débogage avec l’héritage Edge dans Visual Studio Code, voir Microsoft Office [Extension de](debug-with-vs-extension.md)déboguer de Visual Studio Code .

Pour déterminer le navigateur que vous utilisez, consultez Navigateurs utilisés par [les Office les autres.](../concepts/browsers-used-by-office-web-add-ins.md) 

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

> [!NOTE]
> Pour installer une version de Office qui utilise l’ancienne vue web Edge ou pour forcer votre version actuelle de Office à utiliser l’ancienne version de Edge, voir Basculer vers l’ancienne vue [web Edge.](#switch-to-the-edge-legacy-webview)

## <a name="debug-a-task-pane-add-in-using-microsoft-edge-devtools-preview"></a>Déboguer un add-in du volet DevTools à l’Microsoft Edge DevTools Preview

1. Installez la [Microsoft Edge DevTools Preview](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab). (Le mot « Aperçu » est dans le nom pour des raisons historiques. Il n’existe pas de version plus récente.)

   > [!NOTE]
   > Si votre add-in dispose d’une commande de add-in qui exécute une fonction, [celle-ci](../design/add-in-commands.md) s’exécute dans un processus de navigateur masqué que les Microsoft Edge DevTools ne peuvent pas détecter ni attacher, de sorte que la technique décrite dans cet article ne peut pas être utilisée pour déboguer le code dans la fonction.

1. [Chargez une](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) version de version et exécutez le module.
1. Exécutez Microsoft Edge DevTools.
1. Dans les outils, ouvrez l’onglet **Local**. Votre complément est répertorié par son nom. (Seuls les processus en cours d’exécution dans EdgeHTML apparaissent sous l’onglet. L’outil ne peut pas être attaché aux processus en cours d’exécution dans d’autres navigateurs ou vues web, notamment Microsoft Edge (WebView2) et Internet Explorer (Trident).)

   :::image type="content" source="../images/edge-devtools-with-add-in-process.png" alt-text="Capture d’écran de Edge DevTools montrant un processus nommé débogage edge hérité.":::

1. Sélectionnez le nom du module pour l’ouvrir dans les outils.
1. Ouvrez l’onglet **Débogueur**.
1. Ouvrez le fichier que vous souhaitez déboguer en suivant les étapes ci-après.

   1. Dans la barre des tâches du débogger, **sélectionnez Afficher la recherche dans les fichiers.** Cette opération ouvre une fenêtre de recherche.
   1. Entrez une ligne de code à partir du fichier que vous souhaitez déboguer dans la zone de recherche. Il doit s’agir d’un fichier qui n’est probablement pas dans un autre fichier.
   1. Sélectionnez le bouton Actualiser.
   1. Dans les résultats de la recherche, sélectionnez la ligne pour ouvrir le fichier de code dans le volet au-dessus des résultats de la recherche.

   :::image type="content" source="../images/open-file-in-edge-devtools.png" alt-text="Capture d’écran de l’onglet débogage Edge DevTools avec 4 composants étiquetés A à D.":::

1. Pour définir un point d’arrêt, sélectionnez la ligne dans le fichier de code. Le point d’arrêt est inscrit dans le volet **Pile des** appels (en bas à droite). Il peut également y avoir un point rouge par ligne dans le fichier de code, mais cela n’apparaît pas de manière fiable.
1. Exécutez les fonctions dans le complément, si nécessaire, afin de déclencher le point d’arrêt.

> [!TIP]
> Pour plus d’informations sur l’utilisation des outils, [voir Microsoft Edge (EdgeHTML) Developer Tools](/archive/microsoft-edge/legacy/developer/devtools-guide/).

## <a name="debug-a-dialog-in-an-add-in"></a>Débogage d’une boîte de dialogue dans un add-in

Si votre application utilise l’API de boîte de dialogue Office, la boîte de dialogue s’exécute dans un processus distinct du volet Des tâches (le cas besoin) et les outils doivent s’attacher à ce processus. Procédez comme suit.

1. Exécutez le module et les outils.
1. Ouvrez la boîte de dialogue, puis sélectionnez **le bouton Actualiser** dans les outils. Le processus de boîte de dialogue s’affiche. Son nom provient de `<title>` l’élément dans le fichier HTML qui est ouvert dans la boîte de dialogue.
1. Sélectionnez le processus pour l’ouvrir et déboguer comme décrit dans la section Déboguer un add-in du volet Des tâches à l’aide de [Microsoft Edge DevTools Preview](#debug-a-task-pane-add-in-using-microsoft-edge-devtools-preview).

   :::image type="content" source="../images/edge-devtools-with-add-in-and-dialog-processes.png" alt-text="Capture d’écran de Edge DevTools montrant un processus nommé Ma boîte de dialogue.":::

## <a name="switch-to-the-edge-legacy-webview"></a>Basculer vers la vue web edge héritée

Il existe deux façons de changer le mode web edge hérité. Vous pouvez exécuter une commande simple dans une invite de commandes ou installer une version de Office qui utilise Edge Legacy par défaut. Nous vous recommandons la première méthode. Mais vous devez utiliser le deuxième scénario dans les scénarios suivants.

- Votre projet a été développé avec Visual Studio et IIS. Il n’est pas node.js base.
- Vous souhaitez être absolument robuste dans vos tests.
- Si, pour une raison quelconque, l’outil de ligne de commande ne fonctionne pas.

### <a name="switch-via-the-command-line"></a>Basculer via la ligne de commande

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-edge-legacy"></a>Installer une version de Office qui utilise l’ancienne version de Edge

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]
