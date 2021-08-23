---
title: Prise en charge d’Internet Explorer 11
description: Découvrez comment prendre en charge Internet Explorer 11 et ES5 Javascript dans votre add-in.
ms.date: 08/13/2021
localization_priority: Normal
ms.openlocfilehash: dea458cbabb71e23432db8cb6eb3dfcddc6e1bac
ms.sourcegitcommit: bc6203dd8f21d1c375039c5ee8f1388ede9be93b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/18/2021
ms.locfileid: "58382941"
---
# <a name="support-internet-explorer-11"></a>Prise en charge d’Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer toujours utilisé dans les Office de recherche**
>
> Microsoft termine la prise en charge d’Internet Explorer, mais cela n’a pas d’incidence significative sur Office des modules. Certaines combinaisons de plateformes et de versions Office, y compris toutes les versions à achat unique jusqu’à Office 2019, continueront d’utiliser le contrôle webview qui est livré avec Internet Explorer 11 pour héberger des applications, comme expliqué dans les [navigateurs](../concepts/browsers-used-by-office-web-add-ins.md)utilisés par les applications Office . En outre, la prise en charge de ces combinaisons, et donc d’Internet Explorer, est toujours requise pour les applications soumises à [AppSource.](/office/dev/store/submit-to-appsource-via-partner-center) Deux choses *changent* :
>
> - AppSource ne teste plus les Office sur le Web l’aide d’Internet Explorer en tant que navigateur. Toutefois, AppSource teste toujours les combinaisons de plateforme et de Office *de bureau* qui utilisent Internet Explorer.
> - [L Script Lab ne prend](../overview/explore-with-script-lab.md) plus en charge Internet Explorer.

Office Les add-ins sont des applications web qui sont affichées dans des IFrames lors de l’exécution sur Office sur le Web. Office Les macros sont affichées à l’aide de contrôles de navigateur incorporés lors de l’exécution dans Office sur Windows ou Office sur Mac. Les contrôles de navigateur incorporés sont fournis par le système d’exploitation ou par un navigateur installé sur l’ordinateur de l’utilisateur.

Si vous envisagez de commercialiser votre application via AppSource ou si vous prévoyez de prendre en charge des versions antérieures de Windows et Office, votre application doit fonctionner dans le contrôle de navigateur in incorporer basé sur Internet Explorer 11 (IE11). Pour plus d’informations sur les combinaisons de Windows et Office utiliser le contrôle de navigateur internet explorer 11, voir Navigateurs utilisés par les Office de [recherche.](../concepts/browsers-used-by-office-web-add-ins.md)

> [!IMPORTANT]
> Internet Explorer 11 ne prend pas en charge certaines fonctionnalités HTML5 telles que les médias, l’enregistrement et l’emplacement. Si votre add-in doit prendre en charge Internet Explorer 11, vous ne pouvez pas utiliser ces fonctionnalités.

Internet Explorer 11 ne prend pas en charge les versions JavaScript ultérieures à ES5. Si vous souhaitez utiliser la syntaxe et les fonctionnalités d’ECMAScript 2015 ou ultérieure, ou TypeScript, vous disposez de deux options, comme décrit dans cet article. Vous pouvez également combiner ces deux techniques.

## <a name="use-a-transpiler"></a>Utiliser un transpiler

Vous pouvez écrire votre code dans TypeScript ou javaScript moderne, puis le transpiler au moment de la build dans JAVAScript ES5. Les fichiers ES5 qui en résultent sont les fichiers que vous téléchargez dans l’application web de votre application.

Il existe deux transpilers populaires. Les deux peuvent fonctionner avec des fichiers sources qui sont TypeScript ou JavaScript post-ES5. Ils fonctionnent également avec React fichiers (.jsx et .tsx).

- [élie](https://babeljs.io/)
- [tsc](https://www.typescriptlang.org/index.html)

Consultez la documentation de l’un d’eux pour plus d’informations sur l’installation et la configuration du transpiler dans votre projet de add-in. Nous vous recommandons d’utiliser un task runner, tel que [Grunt](https://gruntjs.com/) ou [WebPack,](https://webpack.js.org/) pour automatiser la transpilation. Pour obtenir un exemple de add-in qui utilise tsc, voir Office [de microsoft Graph React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/3ce0e1b74152dbbe8306a091696bc4455c04c0a1/Samples/auth/Office-Add-in-Microsoft-Graph-React). Pour obtenir un exemple qui utilise la bibliothèque d’applications, voir [Offline Stockage Add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/3ce0e1b74152dbbe8306a091696bc4455c04c0a1/Samples/Excel.OfflineStorageAddin).

> [!NOTE]
> Si vous utilisez Visual Studio (pas Visual Studio Code), tsc est probablement plus simple à utiliser. Vous pouvez installer la prise en charge avec un package nuget. Pour plus d’informations, [voir JavaScript et TypeScript dans Visual Studio 2019](/visualstudio/javascript/javascript-in-vs-2019). Pour utiliser l’outil Visual Studio, créez un script de build ou utilisez l’Explorateur de séquenceur de tâches dans Visual Studio avec des outils tels que l’outil [WebPack Task Runner](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) ou [NPM Task Runner.](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner)

## <a name="use-a-polyfill"></a>Utiliser un polyfill

Un [polyfill est](https://en.wikipedia.org/wiki/Polyfill_(programming)) une version antérieure de JavaScript qui duplique les fonctionnalités des versions plus récentes de JavaScript. Le polyfill fonctionne avec dans les navigateurs qui ne sont pas en charge les versions ultérieures de JavaScript. Par exemple, la méthode de chaîne ne faisait pas partie de la version ES5 de JavaScript et ne s’exécutera donc pas dans `startsWith` Internet Explorer 11. Il existe des bibliothèques polyfill, écrites dans ES5, qui définissent et implémentent une `startsWith` méthode. Nous vous recommandons la bibliothèque de polyfill [core-js.](https://github.com/zloirock/core-js)

Pour utiliser une bibliothèque polyfill, chargez-la comme n’importe quel autre fichier ou module JavaScript. Par exemple, vous pouvez utiliser une balise dans le fichier HTML de la page d’accueil du add-in (par exemple), ou vous pouvez utiliser une instruction dans un fichier `<script>` `<script src="/js/core-js.js"></script>` `import` JavaScript (par exemple, `import 'core-js';` ). Lorsque le moteur JavaScript voit une méthode comme , il recherche d’abord s’il existe une méthode de ce nom `startsWith` intégrée dans la langue. Si c’est le cas, il appellera la méthode native. Si la méthode n’est pas intégrée et uniquement si elle n’est pas intégrée, le moteur recherche la méthode dans tous les fichiers chargés. Ainsi, la version polyfilled n’est pas utilisée dans les navigateurs qui la prise en charge de la version native.

L’importation de l’intégralité de la bibliothèque core-js importe toutes les fonctionnalités core-js. Vous pouvez également importer uniquement les polyfills dont votre Office a besoin. Pour obtenir des instructions sur la façon de faire, voir [les API CommonJS.](https://github.com/zloirock/core-js#commonjs-api) La bibliothèque Core-js dispose de la plupart des polyfills dont vous avez besoin. Il existe quelques exceptions détaillées dans la section [Polyfills manquants](https://github.com/zloirock/core-js#missing-polyfills) de la documentation core-js. Par exemple, il ne prend pas en charge, mais vous pouvez utiliser le `fetch` polyfill [d’extraction.](https://github.com/github/fetch)

Pour obtenir un exemple de core.js, consultez l’exemple de [add-in Word Angular2 StyleChecker.](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)

## <a name="testing-an-add-in-on-internet-explorer"></a>Test d’un add-in sur Internet Explorer

Voir [les tests d’Internet Explorer 11.](../testing/ie-11-testing.md)

## <a name="additional-resources"></a>Ressources supplémentaires

- [Tableau de compatibilité ECMAScript 6](https://kangax.github.io/compat-table/es6/)
- [Puis-je utiliser... Tables de prise en charge pour HTML5, CSS3, etc.](https://caniuse.com/)
