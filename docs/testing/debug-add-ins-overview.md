---
title: Débogage des compléments Office
description: Recherchez les Office de débogage des modules pour votre environnement de développement.
ms.date: 12/02/2021
ms.localizationpriority: high
ms.openlocfilehash: aa98bda4de1786f58b730b2375e5586d2cb8b0ad
ms.sourcegitcommit: 33824aa3995a2e0bcc6d8e67ada46f296c224642
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/12/2022
ms.locfileid: "61766097"
---
# <a name="overview-of-debugging-office-add-ins"></a>Vue d’ensemble du débogage Office des modules

Le débogage Office les applications est essentiellement identique au débogage de n’importe quelle application web. Toutefois, un seul ensemble d’outils ne fonctionne pas pour tous les développeurs de modules. Cela est dû au fait que les compléments peuvent être développés sur différents systèmes d’exploitation et s’exécuter sur plusieurs plateformes. Cet article vous aide à trouver les instructions de débogage détaillées pour votre environnement de développement.

> [!TIP]
> Cet article traite du débogage dans le sens étroit de la définition de points d’arrêt et du code pas à pas. Pour obtenir des conseils sur les tests et la résolution des problèmes, commencez par tester les Office et résoudre les [erreurs](test-debug-office-add-ins.md) de développement avec Office les [autres.](troubleshoot-development-errors.md)

> [!NOTE]
> Bien que vous devrez *tester* votre complément sur toutes les plateformes que vous souhaitez prendre en charge, vous n’aurez que très rarement besoin de *déboguer* sur un environnement différent de votre ordinateur de développement. Pour cette raison, cet article utilise « votre ordinateur de développement » et « votre environnement de développement » pour faire référence à l’environnement sur lequel vous déboguer. Si un problème dans le code se produit uniquement sur une plateforme autre que celle de votre ordinateur de développement et que vous devez définir des points d’arrêt ou un code pas à pas pour le résoudre, l’environnement sur lequel vous déboguer n’est pas littéralement votre environnement de développement.

## <a name="server-side-or-client-side"></a>Côté serveur ou côté client ?

Le débogage du code côté serveur d’un Office est identique au débogage côté serveur d’une application web. Consultez les instructions de débogage pour votre IDE ou d’autres outils. Voici quelques exemples pour certains des outils les plus populaires.

- [Déboguer ASP.NET ou ASP.NET Core applications dans Visual Studio](/visualstudio/debugger/how-to-enable-debugging-for-aspnet-applications)
- [Débogage Express](https://expressjs.com/en/guide/debugging.html)
- [Node.js de débogage](https://nodejs.org/en/docs/guides/debugging-getting-started/)
- [Node.js débogage dans VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)
- [Débogage Webpack](https://webpack.js.org/contribute/debugging/)

Le reste de cet article ne concerne que le débogage du JavaScript côté client (qui peut être transposé de TypeScript).

Pour trouver des conseils pour le débogage du code côté client, la première variable est le système d’exploitation de votre ordinateur de développement.

- [Windows](#debug-on-windows)
- [Mac](#debug-on-mac)
- [Linux ou une autre variante Unix](#debug-on-linux)

## <a name="debug-on-windows"></a>Débogage sur Windows

L’exemple suivant fournit des instructions générales sur le débogage sur Windows. Il existe des instructions spéciales pour le débogage de fonctions personnalisées sans interface utilisateur dans des Excel et des Outlook. Consultez [les cas particuliers Windows](#special-cases-in-windows) plus loin dans cette section. Le débogage sur Windows dépend de votre IDE :

- **Visual Studio**: Déboguer à l’aide du déboguer interne Afficher [Debug Office Add-ins in Visual Studio](../develop/debug-office-add-ins-in-visual-studio.md)
- **Visual Studio Code**: Déboguer à l’aide [de l’extension de déboguer de Visual Studio Code](debug-with-vs-extension.md).
- **Tout autre IDE** (ou que vous ne souhaitez pas déboguer à l’intérieur de votre IDE) : utilisez les outils de développement associés au runtime du navigateur que les compléments utilisent sur votre ordinateur de développement. Consultez l’une des rubriques suivantes :

    - [Déboguer des compléments à l’aide des outils de développement pour Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
    - [Déboguer des compléments à l’aide des outils de développement pour la version héritée Edge](debug-add-ins-using-devtools-edge-legacy.md)
    - [Déboguer des compléments à l’aide des Outils de développement dans Microsoft Edge (basés sur Chromium)](debug-add-ins-using-devtools-edge-chromium.md)

Pour plus d’informations sur le runtime du navigateur utilisé, voir Navigateurs utilisés par [les Office de recherche.](../concepts/browsers-used-by-office-web-add-ins.md)

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

### <a name="special-cases-in-windows"></a>Cas particuliers dans Windows

Pour déboguer des fonctions personnalisées sans interface utilisateur Windows, voir débogage de fonctions personnalisées sans interface [utilisateur.](../excel/custom-functions-debugging.md)

Pour déboguer des compléments basés sur des événements dans Outlook, consultez [Déboguer votre Outlook d’événement.](../outlook/debug-autolaunch.md) Le processus nécessite une Visual Studio Code.

## <a name="debug-on-mac"></a>Débogage sur Mac

L’exemple suivant fournit des instructions générales sur le débogage sur Mac. Il existe des instructions spéciales pour le débogage de fonctions personnalisées sans interface utilisateur dans Excel. Afficher [les cas particuliers dans Mac](#special-cases-in-mac) plus loin dans cette section

- Si vous utilisez Visual Studio Code, déboguer à l’aide de [l’extension de](debug-with-vs-extension.md)déboguer de Visual Studio Code .
- Pour tout autre IDE, utilisez l’inspecteur web Safari. Les instructions sont dans [Déboguer des Office sur un Mac.](debug-office-add-ins-on-ipad-and-mac.md)

### <a name="special-cases-in-mac"></a>Cas particuliers dans Mac

Pour déboguer des fonctions personnalisées sans interface utilisateur sur Mac, afficher [débogage des fonctions personnalisées sans interface utilisateur.](../excel/custom-functions-debugging.md)

## <a name="debug-on-linux"></a>Débogage sur Linux

Il n’existe aucune version de bureau de Office pour [ Linux. Vous devrez donc recharger une version test du Office sur le Web](sideload-office-add-ins-for-testing.md) pour le tester et le déboguer. Les conseils de [débogage se trouve dans les compléments de débogage Office sur le Web](debug-add-ins-in-office-online.md).

> [!NOTE]
> Nous vous déconseillons de développer des compléments Office sur un ordinateur Linux, sauf dans le cas inhabituel où vous pouvez vous assurer que tous les utilisateurs du module accéderont au module par le biais de Office sur le Web à partir d’un ordinateur Linux.
