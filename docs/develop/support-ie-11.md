---
title: Prise en charge d’Internet Explorer 11
description: Découvrez comment prendre en charge Internet Explorer 11 et ES5 Javascript dans votre add-in.
ms.date: 10/22/2021
ms.localizationpriority: medium
ms.openlocfilehash: d2a504a6e030e6cf8d06c766cb500d6c11710ea9
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744235"
---
# <a name="support-internet-explorer-11"></a>Prise en charge d’Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer toujours utilisé dans les Office de recherche**
>
> Microsoft termine la prise en charge d’Internet Explorer, mais cela n’a pas d’incidence significative sur Office des modules. Certaines combinaisons de plateformes et de versions Office, y compris les versions à achat unique jusqu’à Office 2019, continueront d’utiliser le contrôle webview qui est livré avec Internet Explorer 11 pour héberger des applications, comme expliqué dans les [navigateurs](../concepts/browsers-used-by-office-web-add-ins.md) utilisés par les applications Office. En outre, la prise en charge de ces combinaisons, et donc d’Internet Explorer, est toujours requise pour les applications soumises à [AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Deux choses *changent* :
>
> - Office sur le Web ne s’ouvre plus dans Internet Explorer. Par conséquent, AppSource ne teste plus les Office sur le Web à l’aide d’Internet Explorer comme navigateur. Toutefois, AppSource teste toujours les combinaisons de plateforme et de Office *de bureau* qui utilisent Internet Explorer.
> - L [Script Lab ne prend](../overview/explore-with-script-lab.md) plus en charge Internet Explorer.

Office sont des applications web qui sont affichées dans des IFrames lors de l’exécution sur Office sur le Web. Office les macros sont affichées à l’aide de contrôles de navigateur incorporés lors de l’exécution dans Office sur Windows ou Office sur Mac. Les contrôles de navigateur incorporés sont fournis par le système d’exploitation ou par un navigateur installé sur l’ordinateur de l’utilisateur.

Si vous envisagez de commercialiser votre add-in via AppSource ou si vous prévoyez de prendre en charge des versions antérieures de Windows et Office, votre application doit fonctionner dans le contrôle de navigateur in incorporer basé sur Internet Explorer 11 (IE11). Pour plus d’informations sur les combinaisons de Windows et Office utiliser le contrôle de navigateur internet explorer 11, voir [Navigateurs](../concepts/browsers-used-by-office-web-add-ins.md) utilisés par les Office de recherche.

> [!IMPORTANT]
> Internet Explorer 11 ne prend pas en charge certaines fonctionnalités HTML5 telles que les médias, l’enregistrement et l’emplacement. Si votre add-in doit prendre en charge Internet Explorer 11, vous devez le concevoir afin d’éviter ces fonctionnalités non prise en charge ou bien il doit détecter quand Internet Explorer est utilisé et offrir une autre expérience qui n’utilise pas les fonctionnalités non prise en charge. Pour plus d’informations, voir [Determine at runtime if the add-in is running in Internet Explorer](#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

## <a name="support-for-recent-versions-of-javascript"></a>Prise en charge des versions récentes de JavaScript

Internet Explorer 11 ne prend pas en charge les versions JavaScript ultérieures à ES5. Si vous souhaitez utiliser la syntaxe et les fonctionnalités d’ECMAScript 2015 ou ultérieure, ou TypeScript, vous disposez de deux options comme décrit dans cet article. Vous pouvez également combiner ces deux techniques.

### <a name="use-a-transpiler"></a>Utiliser un transpiler

Vous pouvez écrire votre code dans TypeScript ou javaScript moderne, puis le transpiler au moment de la build dans JAVAScript ES5. Les fichiers ES5 qui en résultent sont les fichiers que vous téléchargez dans l’application web de votre application.

Il existe deux transpilers populaires. Les deux peuvent fonctionner avec des fichiers sources qui sont TypeScript ou JavaScript post-ES5. Ils fonctionnent également avec React fichiers (.jsx et .tsx).

- [élie](https://babeljs.io/)
- [tsc](https://www.typescriptlang.org/index.html)

Consultez la documentation de l’un d’eux pour plus d’informations sur l’installation et la configuration du transpiler dans votre projet de add-in. Nous vous recommandons d’utiliser un task runner, tel que [Grunt](https://gruntjs.com/) ou [WebPack](https://webpack.js.org/) , pour automatiser la transpilation. Pour obtenir un exemple de add-in qui utilise tsc, voir Office [de microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React). Pour obtenir un exemple qui utilise l’utilisation de la bibliothèque d’applications, voir [Offline Stockage Add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin).

> [!NOTE]
> Si vous utilisez Visual Studio (pas Visual Studio Code), tsc est probablement plus simple à utiliser. Vous pouvez installer la prise en charge avec un package nuget. Pour plus d’informations, [voir JavaScript et TypeScript dans Visual Studio 2019](/visualstudio/javascript/javascript-in-vs-2019). Pour utiliser l’outil Visual Studio, créez un script de build ou utilisez l’Explorateur de séquenceur de tâches dans Visual Studio avec des outils tels que l’outil [WebPack Task Runner](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) ou [NPM Task Runner](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner).

### <a name="use-a-polyfill"></a>Utiliser un polyfill

Un [polyfill est](https://en.wikipedia.org/wiki/Polyfill_(programming)) une version antérieure de JavaScript qui duplique les fonctionnalités des versions plus récentes de JavaScript. Le polyfill fonctionne avec dans les navigateurs qui ne sont pas en charge les versions ultérieures de JavaScript. Par exemple, la méthode de `startsWith` chaîne ne faisait pas partie de la version ES5 de JavaScript et ne s’exécutera donc pas dans Internet Explorer 11. Il existe des bibliothèques polyfill, écrites dans ES5, qui définissent et implémentent une `startsWith` méthode. Nous vous recommandons [la bibliothèque polyfill core-js](https://github.com/zloirock/core-js) .

Pour utiliser une bibliothèque polyfill, chargez-la comme n’importe quel autre fichier ou module JavaScript. Par exemple, `<script>` vous pouvez utiliser une balise dans le fichier HTML de la page d’accueil du module ( `<script src="/js/core-js.js"></script>`par exemple), `import` ou vous pouvez utiliser une instruction dans un fichier JavaScript (par exemple, `import 'core-js';`). Lorsque le moteur JavaScript voit `startsWith`une méthode comme , il recherche d’abord s’il existe une méthode de ce nom intégrée dans le langage. Si c’est le cas, il appellera la méthode native. Si la méthode n’est pas intégrée et uniquement si elle n’est pas intégrée, le moteur recherche la méthode dans tous les fichiers chargés. Ainsi, la version polyfilled n’est pas utilisée dans les navigateurs qui la prise en charge de la version native.

L’importation de l’intégralité de la bibliothèque core-js importe toutes les fonctionnalités core-js. Vous pouvez également importer uniquement les polyfills dont votre Office a besoin. Pour obtenir des instructions sur la façon de faire, consultez les [API CommonJS](https://github.com/zloirock/core-js#commonjs-api). La bibliothèque Core-js dispose de la plupart des polyfills dont vous avez besoin. Il existe quelques exceptions détaillées dans la section [Polyfills manquantes](https://github.com/zloirock/core-js#missing-polyfills) de la documentation core-js. Par exemple, il ne prend pas en charge `fetch`, mais vous pouvez utiliser le polyfill [d’extraction](https://github.com/github/fetch) .

Pour obtenir un exemple de core.js, consultez la classe [Word Add-in Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## <a name="determine-at-runtime-if-the-add-in-is-running-in-internet-explorer"></a>Déterminer au moment de l’exécution si le module est en cours d’exécution dans Internet Explorer

Votre add-in peut découvrir s’il s’exécute dans Internet Explorer en lisant la [propriété window.navigator.userAgent](https://developer.mozilla.org/docs/Web/API/Navigator/userAgent) . Cela permet au module de fournir une expérience de remplacement ou d’échouer normalement. Voici un exemple. Notez qu’Internet Explorer envoie une chaîne commençant par « Trident » comme valeur de userAgent.

```javascript
if (navigator.userAgent.indexOf("Trident") === -1) {

    // IE is not the browser. Provide a full-featured version of the add-in here.

} else {

    // IE is the browser. So here, do one of the following: 
    //  1. Provide an alternate experience that does not use any of the HTML5
    //     features that are not supported in IE.
    //  2. Enable the add-in to gracefully fail by putting a message in the UI that
    //     says something similar to: 
    //      "This add-in won't run in your version of Office. Please upgrade to 
    //      either one-time purchase Office 2021 or to a Microsoft 365 account."          

}
```

> [!IMPORTANT]
> La lecture de la propriété n’est généralement pas `userAgent` une bonne pratique. Assurez-vous que vous êtes familiarisé avec l’article, détection du navigateur à l’aide de [l’agent](https://developer.mozilla.org/en-US/docs/Web/HTTP/Browser_detection_using_the_user_agent) utilisateur, y compris les recommandations et les alternatives à la lecture `userAgent`. En particulier, si vous utilisez l’option 1 `else` dans la clause ci-dessus, envisagez d’utiliser la détection de fonctionnalités au lieu de tester l’agent utilisateur.
>
> Depuis le 30 septembre 2021, le texte de la section Quelle partie de [l’agent](https://developer.mozilla.org/en-US/docs/Web/HTTP/Browser_detection_using_the_user_agent#which_part_of_the_user_agent_contains_the_information_you_are_looking_for) utilisateur contient les informations que vous recherchez ? date d’avant la publication d’Internet Explorer 11. Il est toujours généralement précis et les *tableaux* de la section de la version anglaise de l’article sont à jour. De même, le texte et, dans la plupart des cas, les tableaux, dans les versions non anglaises de l’article sont hors de la date.

## <a name="test-an-add-in-on-internet-explorer"></a>Tester un add-in sur Internet Explorer

Voir [les tests d’Internet Explorer 11](../testing/ie-11-testing.md).

## <a name="additional-resources"></a>Ressources supplémentaires

- [Tableau de compatibilité ECMAScript 6](https://kangax.github.io/compat-table/es6/)
- [Puis-je utiliser... Tables de prise en charge pour HTML5, CSS3, etc.](https://caniuse.com/)
