---
title: Test d’Internet Explorer 11
description: Testez votre Office sur Internet Explorer 11.
ms.date: 11/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8932545aa692073babeddb6ab22a213466a7c2ba
ms.sourcegitcommit: a3debae780126e03a1b566efdec4d8be83e405b8
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/03/2021
ms.locfileid: "60809036"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a>Tester votre Office sur Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer toujours utilisé dans les Office de recherche**
>
> Microsoft termine la prise en charge d’Internet Explorer, mais cela n’a pas d’incidence significative sur Office des modules. Certaines combinaisons de plateformes et de versions Office, y compris les versions à achat unique jusqu’à Office 2019, continueront d’utiliser le contrôle webview qui est livré avec Internet Explorer 11 pour héberger des applications, comme expliqué dans les [navigateurs](../concepts/browsers-used-by-office-web-add-ins.md)utilisés par les Office. En outre, la prise en charge de ces combinaisons, et donc d’Internet Explorer, est toujours requise pour les applications soumises à [AppSource.](/office/dev/store/submit-to-appsource-via-partner-center) Deux choses *changent* :
>
> - Office sur le Web ne s’ouvre plus dans Internet Explorer. Par conséquent, AppSource ne teste plus les Office sur le Web à l’aide d’Internet Explorer en tant que navigateur. Toutefois, AppSource teste toujours les combinaisons de plateforme et de Office *de bureau* qui utilisent Internet Explorer.
> - [L Script Lab ne prend](../overview/explore-with-script-lab.md) plus en charge Internet Explorer.

Si vous envisagez de commercialiser votre application via AppSource ou si vous prévoyez de prendre en charge des versions antérieures de Windows et Office, votre application doit fonctionner dans le contrôle de navigateur in incorporer basé sur Internet Explorer 11 (IE11). Vous pouvez utiliser une ligne de commande pour passer de runtimes plus modernes utilisés par les modules de mise à l’essai à Internet Explorer 11 pour ce test. Pour plus d’informations sur les versions de Windows et Office utiliser le contrôle d’affichage web Internet Explorer 11, voir Navigateurs utilisés par les Office des [applications.](../concepts/browsers-used-by-office-web-add-ins.md)

> [!IMPORTANT]
> Internet Explorer 11 ne prend pas en charge les versions de JavaScript ultérieures à la version ES5. Si vous souhaitez utiliser la syntaxe et les fonctionnalités d’ECMAScript 2015 ou une ultérieure, vous disposez de deux options :
>
> - Écrivez votre code dans ECMAScript 2015 (également appelé ES6) ou version ultérieure JavaScript, ou dans TypeScript, puis compilez votre code en JavaScript ES5 à l’aide d’un compilateur tel que [celui-ci ou](https://babeljs.io/) [tsc.](https://www.typescriptlang.org/index.html)
> - Écrivez en JavaScript ECMAScript 2015 ou version ultérieure, mais chargez également une [bibliothèque polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) telle que [core-js](https://github.com/zloirock/core-js) qui permet à IE d’exécuter votre code.
>
> Pour plus d’informations sur ces options, voir [Support Internet Explorer 11](../develop/support-ie-11.md).
>
> Par ailleurs, Internet Explorer 11 ne prend pas en charge certaines fonctionnalités HTML5 telles que les éléments multimédias, l’enregistrement et l’emplacement. Pour plus d’informations, voir Déterminer au moment de l’exécution si le module est en cours d’exécution [dans Internet Explorer.](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)

> [!NOTE]
> Office sur le Web ne peut pas être ouvert dans Internet Explorer 11, vous ne pouvez pas (et n’avez pas besoin de) tester votre module sur Office sur le Web avec Internet Explorer.

## <a name="switch-to-the-internet-explorer-11-webview"></a>Basculer vers la vue web d’Internet Explorer 11

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

Il existe deux façons de basculer le mode web d’Internet Explorer. Vous pouvez exécuter une commande simple dans une invite de commandes ou installer une version de Office qui utilise Internet Explorer par défaut. Nous vous recommandons la première méthode. Mais vous devez utiliser le deuxième scénario dans les scénarios suivants.

- Votre projet a été développé avec Visual Studio et IIS. Il n’est pas node.js base.
- Vous souhaitez être absolument robuste dans vos tests.
- Si, pour une raison quelconque, l’outil de ligne de commande ne fonctionne pas.

### <a name="switch-via-the-command-line"></a>Basculer via la ligne de commande

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-internet-explorer"></a>Installer une version de Office qui utilise Internet Explorer

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## <a name="see-also"></a>Voir aussi

* [Test et débogage de compléments Office](test-debug-office-add-ins.md)
* [Chargement de la version test des compléments Office](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [Déboguer des compléments à l’aide des outils de développement pour Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
* [Attacher un débogueur à partir du volet Office](attach-debugger-from-task-pane.md)
