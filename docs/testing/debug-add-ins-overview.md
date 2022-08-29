---
title: Débogage des compléments Office
description: Recherchez les conseils de débogage des compléments Office pour votre environnement de développement.
ms.date: 07/11/2022
ms.localizationpriority: high
ms.openlocfilehash: f23e55b2d3ceb84e32365ffbbcb9efafedfebcfc
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423271"
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

## <a name="special-cases"></a>Cas particuliers

Dans certains cas particuliers, le processus de débogage diffère de la normale pour une combinaison donnée de plateforme, d’application Office et d’environnement de développement. Si vous déboguez l’un de ces cas spéciaux, utilisez les liens de cette section pour trouver les conseils appropriés. Sinon, continuez à [Instructions générales](#general-guidance).

- **Débogage la fonction `Office.initialize` ou `Office.onReady`** : [Déboguer les fonctions initialize et onReady](debug-initialize-onready.md).
- **Debugging d’une fonction personnalisée Excel _dans un_ runtime non-partagé** : [Débogage de fonctions personnalisées dans un runtime non-partagé](../excel/custom-functions-debugging.md).
- **Débogage d'une [commande de fonction](../design/add-in-commands.md#types-of-add-in-commands)dans un __ runtime non-partagée** : 
    - Compléments Outlook sur un ordinateur de développement Windows : [commandes de fonction de débogage dans les compléments Outlook](../outlook/debug-ui-less.md) 
    - Autres compléments d’application Office ou Outlook sur un ordinateur de développement Mac : [Déboguez une commande de fonction avec un runtime non partagé](debug-function-command.md).
- **Débogage d'un complément Outlook basé sur des événements** : [Déboguez votre complément Outlook basé sur des événements](../outlook/debug-autolaunch.md). 
 
## <a name="general-guidance"></a>Directives générales

Pour trouver des conseils pour le débogage du code côté client, la première variable est le système d’exploitation de votre ordinateur de développement.

- [Windows](#debug-on-windows)
- [Mac](#debug-on-mac)
- [Linux ou une autre variante Unix](#debug-on-linux)

### <a name="debug-on-windows"></a>Débogage sur Windows

L’exemple suivant fournit des instructions générales sur le débogage sur Windows. Le débogage sur Windows dépend de votre IDE.

- **Visual Studio** : Déboguer à l’aide des outils F12 du navigateur. Afficher [Debug Office Add-ins in Visual Studio](../develop/debug-office-add-ins-in-visual-studio.md)
- **Visual Studio Code**: Déboguer à l’aide [de l’extension de déboguer de Visual Studio Code](debug-with-vs-extension.md).
- **Tout autre** IDE (ou vous ne voulez pas déboguer dans votre IDE) : Utilisez les outils de développement qui sont associés au moteur d'exécution du navigateur que les compléments utilisent sur votre ordinateur de développement. Voir l'un des documents suivants :

    - [Déboguer des compléments à l’aide des outils de développement pour Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
    - [Déboguer des compléments à l’aide des outils de développement pour la version héritée Edge](debug-add-ins-using-devtools-edge-legacy.md)
    - [Déboguer des compléments à l’aide des Outils de développement dans Microsoft Edge (basés sur Chromium)](debug-add-ins-using-devtools-edge-chromium.md)

Pour plus d’informations sur le runtime utilisé, consultez [Navigateurs utilisés par les compléments](../concepts/browsers-used-by-office-web-add-ins.md) Office et [les runtimes dans les compléments Office](runtimes.md).

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

### <a name="debug-on-mac"></a>Débogage sur Mac

L’exemple suivant fournit des instructions générales sur le débogage sur Mac.

- Si vous utilisez Visual Studio Code, déboguer à l’aide de [l’extension de](debug-with-vs-extension.md)déboguer de Visual Studio Code .
- Pour tout autre IDE, utilisez l’inspecteur web Safari. Les instructions sont dans [Déboguer des Office sur un Mac.](debug-office-add-ins-on-ipad-and-mac.md)


### <a name="debug-on-linux"></a>Débogage sur Linux

Il n'existe pas de version de bureau d'Office pour Linux. Vous devrez donc charger [le complément dans Office on the web](sideload-office-add-ins-for-testing.md) pour le tester et le déboguer. Vous trouverez des conseils sur le débogage dans [Debug des compléments dans Office on the web](debug-add-ins-in-office-online.md).

> [!NOTE]
> Nous vous déconseillons de développer des compléments Office sur un ordinateur Linux, sauf dans le cas inhabituel où vous pouvez vous assurer que tous les utilisateurs du module accéderont au module par le biais de Office sur le Web à partir d’un ordinateur Linux.

## <a name="debug-add-ins-in-staging-or-production"></a>Déboguer des compléments en préproduction ou en production

Pour déboguer un complément déjà en préproduction ou en production, attachez un débogueur à partir de l’interface utilisateur du complément. Pour obtenir des instructions, consultez [Attacher un débogueur à partir du volet Office](attach-debugger-from-task-pane.md).

## <a name="see-also"></a>Voir aussi

- [Runtimes dans les compléments Office](runtimes.md)
