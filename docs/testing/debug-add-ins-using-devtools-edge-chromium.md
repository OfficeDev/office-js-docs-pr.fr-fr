---
title: Déboguer des applications à l’aide des outils de développement pour Microsoft Edge WebView2
description: Déboguer des applications à l’aide des outils de développement Microsoft Edge WebView2 (Chromium web).
ms.date: 11/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: d2a8aa41720eebcd15d4cb625d90af1399b87dad
ms.sourcegitcommit: 3d37c42f5e465dac52d231d31717bdbb3bfa0e30
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/10/2021
ms.locfileid: "60923540"
---
# <a name="debug-add-ins-using-developer-tools-in-microsoft-edge-chromium-based"></a>Déboguer des applications à l’aide des outils de développement Microsoft Edge (Chromium base de données)

Cet article montre comment déboguer le code côté client (JavaScript ou TypeScript) de votre add-in lorsque les conditions suivantes sont remplies.

- Vous ne pouvez pas ou ne souhaitez pas déboguer à l’aide des outils intégrés à votre IDE ; ou vous rencontrez un problème qui se produit uniquement lorsque le module est exécuté en dehors de l’IDE.
- Votre ordinateur utilise une combinaison de versions Windows et Office qui utilisent le contrôle WebView Edge (basé sur Chromium), WebView2.

> [!TIP]
> Pour plus d’informations sur le débogage avec Edge WebView2 (basé sur Chromium) à l’intérieur de Visual Studio Code, voir Déboguer des applications sur Windows à l’aide de Visual Studio Code et [Microsoft Edge WebView2 (basé](debug-desktop-using-edge-chromium.md)sur Chromium).

Pour déterminer le navigateur que vous utilisez, consultez Navigateurs utilisés par [les Office les autres.](../concepts/browsers-used-by-office-web-add-ins.md)

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

## <a name="debug-a-task-pane-add-in-using-microsoft-edge-chromium-based-developer-tools"></a>Déboguer un add-in du volet De tâches à l’aide Microsoft Edge de développement Chromium (basés sur Chromium)

> [!NOTE]
> Si votre application dispose d’une commande de [add-in](../design/add-in-commands.md) qui exécute une fonction, celle-ci s’exécute dans un processus de navigateur masqué à partir de qui les outils de développement Microsoft Edge (basés sur Chromium) ne peuvent pas être lancés, de sorte que la technique décrite dans cet article ne peut pas être utilisée pour déboguer du code dans la fonction.

1. [Chargez une](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) version de version et exécutez le module.
1. Exécutez les Microsoft Edge de développement Chromium (basés sur Chromium) à l’aide de l’une des méthodes ci-après :

   - Assurez-vous que le volet Des tâches du add-in a le focus et appuyez sur **Ctrl+Shift+I**.
   - Cliquez avec le bouton droit sur le volet Des tâches pour ouvrir le menu contexto et sélectionnez **Inspecter,** ou ouvrez le [menu](../design/task-pane-add-ins.md#personality-menu) De personnalité et sélectionnez **Attacher le débogger.**

1. Ouvrez **l’onglet Sources.**
1. Ouvrez le fichier que vous souhaitez déboguer en suivant les étapes ci-après.

   1. À l’extrême droite de la barre de menus supérieure de l’outil, sélectionnez le bouton **...** puis sélectionnez **Rechercher.**
   1. Entrez une ligne de code à partir du fichier que vous souhaitez déboguer dans la zone de recherche. Il doit s’agir d’un fichier qui n’est probablement pas dans un autre fichier.
   1. Sélectionnez le bouton Actualiser.
   1. Dans les résultats de la recherche, sélectionnez la ligne pour ouvrir le fichier de code dans le volet au-dessus des résultats de la recherche.

   :::image type="content" source="../images/open-file-in-edge-chromium-devtools.png" alt-text="Capture d’écran de l Chromium onglet source des outils de développement Edge avec 4 composants étiquetés A à D.":::

1. Pour définir un point d’arrêt, sélectionnez le numéro de ligne de la ligne dans le fichier de code. Un point rouge apparaît par la ligne dans le fichier de code. Dans la fenêtre du débogger à droite, le point d’arrêt est inscrit dans la liste des **points** d’arrêt.
1. Exécutez les fonctions dans le complément, si nécessaire, afin de déclencher le point d’arrêt.

> [!TIP]
> Pour plus d’informations sur l’utilisation des outils, [voir Microsoft Edge outils de développement.](/microsoft-edge/devtools-guide-chromium/)

## <a name="debug-a-dialog-in-an-add-in"></a>Débogage d’une boîte de dialogue dans un add-in

Si votre application utilise l’API de boîte de dialogue Office, la boîte de dialogue s’exécute dans un processus distinct du volet Des tâches (le cas besoin) et l’outil doit être démarré à partir de ce processus distinct. Procédez comme suit.

1. Exécutez le complément.
1. Ouvrez la boîte de dialogue et assurez-vous qu’elle a le focus.
1. Ouvrez Microsoft Edge outils de développement (basés sur Chromium) à l’aide de l’une des méthodes ci-après :

   - Appuyez **sur Ctrl+Shift+I** ou **F12**.
   - Cliquez avec le bouton droit sur la boîte de dialogue pour ouvrir le menu contextnel et sélectionnez **Inspecter.**

1. Utilisez l’outil de la même manière que pour le code dans un volet De tâches. Voir [Déboguer un add-in](#debug-a-task-pane-add-in-using-microsoft-edge-chromium-based-developer-tools) du volet Des tâches à l’aide Microsoft Edge (basés sur Chromium) plus tôt dans cet article.
