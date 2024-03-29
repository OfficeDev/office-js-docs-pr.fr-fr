---
title: Prise en charge d’Internet Explorer 11
description: Découvrez comment prendre en charge Internet Explorer 11 et ES5 Javascript dans votre complément.
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: aff6004af4ce28aea865cb34cd34e13e23fb549f
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810273"
---
# <a name="support-internet-explorer-11"></a>Prise en charge d’Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer est toujours utilisé dans les compléments Office**
>
> Certaines combinaisons de plateformes et de versions d’Office, y compris les versions perpétuelles via Office 2019, utilisent toujours le contrôle webview fourni avec Internet Explorer 11 pour héberger les compléments, comme expliqué dans [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md). Nous vous recommandons (mais n’exigez pas) de continuer à prendre en charge ces combinaisons, au moins de manière minimale, en fournissant aux utilisateurs de votre complément un message d’échec approprié lorsque votre complément est lancé dans la vue web Internet Explorer. Gardez ces points supplémentaires à l’esprit :
>
> - Office sur le Web ne s’ouvre plus dans Internet Explorer. Par conséquent, [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) ne teste plus les compléments dans Office sur le Web en utilisant Internet Explorer comme navigateur.
> - AppSource teste toujours les combinaisons de versions de *plateforme* et de bureau Office qui utilisent Internet Explorer, mais elle émet un avertissement uniquement lorsque le complément ne prend pas en charge Internet Explorer ; le complément n’est pas rejeté par AppSource.
> - [L’outil Script Lab](../overview/explore-with-script-lab.md) ne prend plus en charge Internet Explorer.

Les compléments Office sont des applications web qui s’affichent dans des IFrames lors de l’exécution sur Office sur le Web. Les compléments Office sont affichés à l’aide de contrôles de navigateur incorporés lors de l’exécution dans Office sur Windows ou Office sur Mac. Les contrôles de navigateur incorporés sont fournis par le système d’exploitation ou par un navigateur installé sur l’ordinateur de l’utilisateur.

Si vous envisagez de prendre en charge des versions antérieures de Windows et Office, votre complément doit fonctionner dans le contrôle de navigateur incorporable basé sur Internet Explorer 11 (IE11). Pour plus d’informations sur les combinaisons de Windows et d’Office qui utilisent le contrôle de navigateur IE11, voir [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Internet Explorer 11 ne prend pas en charge certaines fonctionnalités HTML5 telles que les médias, l’enregistrement et l’emplacement. Si votre complément doit prendre en charge Internet Explorer 11, vous devez concevoir le complément pour éviter ces fonctionnalités non prises en charge ou le complément doit détecter quand Internet Explorer est utilisé et fournir une autre expérience qui n’utilise pas les fonctionnalités non prises en charge. Pour plus d’informations, consultez [Déterminer au moment de l’exécution si le complément s’exécute dans Internet Explorer](#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

## <a name="support-for-recent-versions-of-javascript"></a>Prise en charge des versions récentes de JavaScript

Internet Explorer 11 ne prend pas en charge les versions javascript ultérieures à ES5. Si vous souhaitez utiliser la syntaxe et les fonctionnalités d’ECMAScript 2015 ou version ultérieure, ou de TypeScript, vous disposez de deux options, comme décrit dans cet article. Vous pouvez également combiner ces deux techniques.

### <a name="use-a-transpiler"></a>Utiliser un transpileur

Vous pouvez écrire votre code en TypeScript ou en JavaScript moderne, puis le transpiler au moment de la génération en JavaScript ES5. Les fichiers ES5 obtenus sont ce que vous chargez dans l’application web de votre complément.

Il existe deux transpileurs populaires. Les deux peuvent fonctionner avec des fichiers sources TypeScript ou JavaScript post-ES5. Ils fonctionnent également avec des fichiers React (.jsx et .tsx).

- [Babel](https://babeljs.io/)
- [Tsc](https://www.typescriptlang.org/index.html)

Pour plus d’informations sur l’installation et la configuration du transpileur dans votre projet de complément, consultez la documentation relative à l’un ou l’autre d’entre eux. Nous vous recommandons d’utiliser un exécuteur de tâches, tel que [Grunt](https://gruntjs.com/) ou [WebPack](https://webpack.js.org/) , pour automatiser la transpilation. Pour obtenir un exemple de complément qui utilise tsc, voir [Complément Office Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React). Pour obtenir un exemple qui utilise babel, consultez [Complément de stockage hors connexion](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin).

> [!NOTE]
> Si vous utilisez Visual Studio (et non Visual Studio Code), tsc est probablement plus facile à utiliser. Vous pouvez installer la prise en charge de celui-ci avec un package nuget. Pour plus d’informations, consultez [JavaScript et TypeScript dans Visual Studio 2019](/visualstudio/javascript/javascript-in-vs-2019). Pour utiliser babel avec Visual Studio, créez un script de build ou utilisez l’Explorateur d’exécuteurs de tâches dans Visual Studio avec des outils tels que [l’exécuteur de tâches WebPack](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) ou [l’exécuteur de tâches NPM](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner).

### <a name="use-a-polyfill"></a>Utiliser un polyfill

Un [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) est une version antérieure de JavaScript qui duplique les fonctionnalités des versions plus récentes de JavaScript. Le polyfill fonctionne avec dans les navigateurs qui ne prennent pas en charge les versions ultérieures de JavaScript. Par exemple, la méthode `startsWith` string ne faisait pas partie de la version ES5 de JavaScript et ne s’exécute donc pas dans Internet Explorer 11. Il existe des bibliothèques polyfill, écrites dans ES5, qui définissent et implémentent une `startsWith` méthode. Nous vous recommandons la bibliothèque de polyfill [core-js](https://github.com/zloirock/core-js) .

Pour utiliser une bibliothèque polyfill, chargez-la comme n’importe quel autre fichier ou module JavaScript. Par exemple, vous pouvez utiliser une `<script>` balise dans le fichier HTML de la page d’accueil du complément (par exemple `<script src="/js/core-js.js"></script>`), ou vous pouvez utiliser une `import` instruction dans un fichier JavaScript (par exemple, `import 'core-js';`). Lorsque le moteur JavaScript voit une méthode telle `startsWith`que , il recherche d’abord s’il existe une méthode de ce nom intégrée dans le langage. Si c’est le cas, il appelle la méthode native. Si, et seulement si, la méthode n’est pas intégrée, le moteur la recherche dans tous les fichiers chargés. Par conséquent, la version polyfilled n’est pas utilisée dans les navigateurs qui prennent en charge la version native.

L’importation de la bibliothèque core-js entière importera toutes les fonctionnalités core-js. Vous pouvez également importer uniquement les polyfills requis par votre complément Office. Pour obtenir des instructions sur la procédure à suivre, consultez [API CommonJS](https://github.com/zloirock/core-js#commonjs-api). La bibliothèque core-js contient la plupart des polyfills dont vous avez besoin. Il existe quelques exceptions détaillées dans la section [Polyfills manquants](https://github.com/zloirock/core-js#missing-polyfills) de la documentation core-js. Par exemple, il ne prend pas en charge `fetch`, mais vous pouvez utiliser le polyfill [d’extraction](https://github.com/github/fetch) .

Pour obtenir un exemple de complément qui utilise core.js, consultez Complément [Word Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## <a name="determine-at-runtime-if-the-add-in-is-running-in-internet-explorer"></a>Déterminer au moment de l’exécution si le complément s’exécute dans Internet Explorer

Votre complément peut découvrir s’il s’exécute dans Internet Explorer en lisant la propriété [window.navigator.userAgent](https://developer.mozilla.org/docs/Web/API/Navigator/userAgent) . Cela permet au complément de fournir une autre expérience ou d’échouer correctement. Voici un exemple. Notez qu’Internet Explorer envoie une chaîne commençant par « Trident » comme valeur de userAgent.

```javascript
if (navigator.userAgent.indexOf("Trident") === -1) {

    // IE is not the browser. Provide a full-featured version of the add-in here.

} else {

    // IE is the browser. So here, do one of the following: 
    //  1. Provide an alternate experience that does not use any of the HTML5
    //     features that are not supported in IE.
    //  2. Enable the add-in to gracefully fail by putting a message in the UI that
    //     says something similar to: 
    //      "This add-in won't run in your version of Office. Please upgrade 
    //      either to perpetual Office 2021 or to a Microsoft 365 account."          

}
```

> [!IMPORTANT]
> Il n’est généralement pas recommandé de lire la `userAgent` propriété. Veillez à connaître l’article Détection du [navigateur à l’aide de l’agent utilisateur](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent), y compris les recommandations et les alternatives à la lecture `userAgent`. En particulier, si vous utilisez l’option 1 dans la `else` clause ci-dessus, envisagez d’utiliser la détection des fonctionnalités au lieu de tester l’agent utilisateur.
>
> Depuis le 30 septembre 2021, le texte de la section [Quelle partie de l’agent utilisateur contient les informations que vous recherchez ?](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent#which_part_of_the_user_agent_contains_the_information_you_are_looking_for) date d’avant la publication d’Internet Explorer 11. Il est toujours généralement exact et les *tableaux* de la section de la version anglaise de l’article sont à jour. De même, le texte, et dans la plupart des cas les tableaux, dans les versions non anglaises de l’article sont obsolètes.

## <a name="test-an-add-in-on-internet-explorer"></a>Tester un complément sur Internet Explorer

Consultez [Test d’Internet Explorer 11](../testing/ie-11-testing.md).

## <a name="additional-resources"></a>Ressources supplémentaires

- [Table de compatibilité ECMAScript 6](https://kangax.github.io/compat-table/es6/)
- [Puis-je utiliser... Tables de prise en charge pour HTML5, CSS3, etc.](https://caniuse.com/)
