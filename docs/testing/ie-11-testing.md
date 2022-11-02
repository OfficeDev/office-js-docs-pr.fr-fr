---
title: Test d’Internet Explorer 11
description: Testez votre complément Office sur Internet Explorer 11.
ms.date: 10/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: f5e962bb615849b4944be2bee3f14006b0c9289e
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810360"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a>Tester votre complément Office sur Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer est toujours utilisé dans les compléments Office**
>
> Certaines combinaisons de plateformes et de versions d’Office, y compris les versions perpétuelles via Office 2019, utilisent toujours le contrôle webview fourni avec Internet Explorer 11 pour héberger les compléments, comme expliqué dans [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md). Nous vous recommandons (mais n’exigez pas) de continuer à prendre en charge ces combinaisons, au moins de manière minimale, en fournissant aux utilisateurs de votre complément un message d’échec approprié lorsque votre complément est lancé dans la vue web Internet Explorer. Gardez ces points supplémentaires à l’esprit :
>
> - Office sur le Web ne s’ouvre plus dans Internet Explorer. Par conséquent, [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) ne teste plus les compléments dans Office sur le Web en utilisant Internet Explorer comme navigateur.
> - AppSource teste toujours les combinaisons de versions de *plateforme* et de bureau Office qui utilisent Internet Explorer, mais elle émet un avertissement uniquement lorsque le complément ne prend pas en charge Internet Explorer ; le complément n’est pas rejeté par AppSource.
> - [L’outil Script Lab](../overview/explore-with-script-lab.md) ne prend plus en charge Internet Explorer.

Si vous envisagez de prendre en charge des versions antérieures de Windows et Office, votre complément doit fonctionner dans le contrôle de navigateur incorporable basé sur Internet Explorer 11 (IE11). Vous pouvez utiliser une ligne de commande pour passer des runtimes plus modernes utilisés par les compléments au runtime Internet Explorer 11 pour ce test. Pour plus d’informations sur les versions de Windows et d’Office qui utilisent le contrôle d’affichage web Internet Explorer 11, voir [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Internet Explorer 11 ne prend pas en charge les versions de JavaScript ultérieures à la version ES5. Si vous souhaitez utiliser la syntaxe et les fonctionnalités d’ECMAScript 2015 ou version ultérieure, vous avez deux options :
>
> - Écrivez votre code dans ECMAScript 2015 (également appelé ES6) ou javaScript ultérieur, ou dans TypeScript, puis compilez votre code dans JavaScript ES5 à l’aide d’un compilateur tel que [babel](https://babeljs.io/) ou [tsc](https://www.typescriptlang.org/index.html).
> - Écrivez dans ECMAScript 2015 ou une version ultérieure de JavaScript, mais chargez également une bibliothèque [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) telle que [core-js](https://github.com/zloirock/core-js) qui permet à Internet Explorer d’exécuter votre code.
>
> Pour plus d’informations sur ces options, consultez [Prise en charge d’Internet Explorer 11](../develop/support-ie-11.md).
>
> Par ailleurs, Internet Explorer 11 ne prend pas en charge certaines fonctionnalités HTML5 telles que les éléments multimédias, l’enregistrement et l’emplacement. Pour plus d’informations, consultez [Déterminer au moment de l’exécution si le complément est en cours d’exécution dans Internet Explorer](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

> [!NOTE]
> - Office sur le Web ne pouvant pas être ouvert dans Internet Explorer 11, vous ne pouvez pas (et n’avez pas besoin de) tester votre complément sur Office sur le Web avec Internet Explorer.
>
> - La Configuration de sécurité renforcée d’Internet Explorer (ESC) doit être désactivée pour que les compléments web Office fonctionnent. Si vous utilisez un ordinateur Windows Server comme votre client lors du développement des compléments, notez qu’ESC est activée par défaut dans Windows Server.

## <a name="switch-to-the-internet-explorer-11-webview"></a>Basculer vers la vue web d’Internet Explorer 11

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

Il existe deux façons de basculer l’affichage web d’Internet Explorer. Vous pouvez exécuter une commande simple dans une invite de commandes ou installer une version d’Office qui utilise Internet Explorer par défaut. Nous vous recommandons la première méthode. Mais vous devez utiliser la seconde dans les scénarios suivants.

- Votre projet a été développé avec Visual Studio et IIS. Il n’est pas basé sur node.js.
- Vous voulez être absolument robuste dans vos tests.
- Vous ne pouvez pas utiliser le canal bêta pour Microsoft 365 sur votre ordinateur de développement.
- Vous développez sur un Mac. 
- Si, pour une raison quelconque, l’outil en ligne de commande ne fonctionne pas.

### <a name="switch-via-the-command-line"></a>Basculer via la ligne de commande

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-internet-explorer"></a>Installer une version d’Office qui utilise Internet Explorer

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## <a name="see-also"></a>Voir aussi

- [Test et débogage de compléments Office](test-debug-office-add-ins.md)
- [Chargement de la version test des compléments Office](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [Déboguer des compléments à l’aide des outils de développement pour Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
- [Attacher un débogueur à partir du volet Office](attach-debugger-from-task-pane.md)
- [Runtimes dans les compléments Office](runtimes.md)