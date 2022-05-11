---
title: Navigateurs utilisés par les compléments Office
description: Indique comment le système d’exploitation et la version d’Office déterminent le navigateur utilisé par les compléments Office.
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5e563c836b48a16f572aca492fa39f33b9661052
ms.sourcegitcommit: fd04b41f513dbe9e623c212c1cbd877ae2285da0
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/11/2022
ms.locfileid: "65313183"
---
# <a name="browsers-used-by-office-add-ins"></a>Navigateurs utilisés par les compléments Office

Office compléments sont des applications web affichées à l’aide d’iFrames lors de l’exécution dans Office sur le Web. Dans Office pour les clients de bureau et mobiles, Office compléments utilisent un contrôle de navigateur incorporé (également appelé webview). Les compléments ont également besoin d’un moteur JavaScript pour exécuter le code JavaScript. Le navigateur incorporé et le moteur sont fournis par un navigateur installé sur l’ordinateur de l’utilisateur.

Le navigateur utilisé dépend de ce qui suit :

- Système d’exploitation de l’ordinateur.
- Que le complément s’exécute dans Office sur le Web, Microsoft 365 ou hors abonnement Office 2013 ou version ultérieure.

> [!IMPORTANT]
> **Internet Explorer toujours utilisé dans les compléments Office**
>
> Certaines combinaisons de plateformes et de versions Office, y compris les versions à achat unique via Office 2019, utilisent toujours le contrôle webview fourni avec Internet Explorer 11 pour héberger des compléments, comme expliqué dans cet article. Nous vous recommandons (mais n’exigez pas) de continuer à prendre en charge ces combinaisons, du moins d’une manière minimale, en fournissant aux utilisateurs de votre complément un message d’échec approprié lorsque votre complément est lancé dans la vue web d’Internet Explorer. Gardez à l’esprit ces points supplémentaires :
>
> - Office sur le Web ne s’ouvre plus dans Internet Explorer. Par conséquent, [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) ne teste plus les compléments dans Office sur le Web à l’aide d’Internet Explorer comme navigateur.
> - AppSource teste toujours les combinaisons *de versions de* plateforme et de bureau Office qui utilisent Internet Explorer, mais elle émet uniquement un avertissement lorsque le complément ne prend pas en charge Internet Explorer ; le complément n’est pas rejeté par AppSource.
> - [L’outil Script Lab](../overview/explore-with-script-lab.md) ne prend plus en charge Internet Explorer.
>
> Pour plus d’informations sur la prise en charge d’Internet Explorer et la configuration d’un message d’échec approprié sur votre complément, consultez [Support Internet Explorer 11](../develop/support-ie-11.md).

Le tableau ci-dessous répertorie le navigateur utilisé selon les plateformes et systèmes d’exploitation.

|Système d’exploitation|Version d’Office|Edge WebView2 (basé sur Chromium) installé ?|Navigateur|
|:-----|:-----|:-----|:-----|
|indifférent|Office sur le web|Non applicable|Navigateur dans lequel Office sur le web est ouvert.<br>(Notez toutefois que Office sur le Web ne s’ouvre pas dans Internet Explorer.<br>Si vous tentez de le faire, Office sur le Web s’ouvre dans Edge.) |
|Mac|indifférent|Non applicable|Safari avec WKWebView|
|iOS|indifférent|Non applicable|Safari avec WKWebView|
|Android|indifférent|Non applicable|Chrome|
|Windows 7, 8.1, 10, 11 | non-abonnement Office 2013 à Office 2019|Peu importe|Internet Explorer 11|
|Windows 10, 11 | Office 2021 sans abonnement ou version ultérieure|Oui|Microsoft Edge <sup>1</sup> avec WebView2 (basé sur Chromium)|
|Windows 7 | Microsoft 365| Peu importe | Internet Explorer 11|
|Windows 8.1,<br>Windows 10 ver.&nbsp;<&nbsp; 1903| Microsoft 365 | Non| Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp; 1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;<&nbsp; 16.0.116292<sup></sup>| Peu importe|Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp; 1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;>=&nbsp; 16.0.11629AND16.0.13530.204242&nbsp;&nbsp;<sup></sup><&nbsp;| Peu importe|Microsoft Edge <sup>1, 3</sup> avec WebView d’origine (EdgeHTML)|
|Windows 10 ver.&nbsp;>=&nbsp; 1903,<br>Fenêtre 11 | Microsoft 365 ver.&nbsp;>=&nbsp; 16.0.13530.204242<sup></sup>| Non |Microsoft Edge <sup>1, 3</sup> avec WebView d’origine (EdgeHTML)|
|Windows 8.1<br>Windows 10,<br>Windows 11| Microsoft 365 ver.&nbsp;>=&nbsp; 16.0.13530.204242<sup></sup>| <sup>Oui4</sup>|  Microsoft Edge <sup>1</sup> avec WebView2 (basé sur Chromium) |

<sup>1</sup> Lorsque Microsoft Edge est utilisé, le Windows Narrateur (parfois appelé « lecteur d’écran ») lit la `<title>` balise dans la page qui s’ouvre dans le volet Office. Si Internet Explorer 11 est utilisé, le Narrateur lit la barre de titre du volet Office, qui provient de la valeur `<DisplayName>` du manifeste du complément.

<sup>2</sup> Pour plus d’informations, consultez la page historique des [mises à jour](/officeupdates/update-history-office365-proplus-by-date) et comment [trouver votre version du client Office et le canal de mise à jour](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).

<sup>3</sup> Si votre complément inclut l’élément `<Runtimes>` dans le manifeste, il n’utilise pas Microsoft Edge avec le WebView d’origine (EdgeHTML). Si les conditions d’utilisation de Microsoft Edge avec WebView2 (basée sur Chromium) sont remplies, le complément utilise ce navigateur. Sinon, il utilise Internet Explorer 11, quelle que soit la version Windows ou Microsoft 365. Pour plus d’informations, voir [Services d’exécution](/javascript/api/manifest/runtimes).

<sup>4</sup> Sur Windows versions antérieures à Windows 11, le contrôle WebView2 doit être installé afin que Office puisse l’incorporer. Il est installé avec Microsoft 365, version 2101 ou ultérieure, et avec un achat unique Office 2021 ou une version ultérieure; mais il n’est pas installé automatiquement avec Microsoft Edge. Si vous disposez d’une version antérieure de Microsoft 365 ou d’un Office d’achat unique, suivez les instructions d’installation du contrôle sur [Microsoft Edge WebView2 / Incorporer du contenu web ... avec Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/). Sur Microsoft 365 builds antérieures à 16.0.14326.xxxxx, vous devez également créer la clé de Registre **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Win32WebView2** et définir sa valeur `dword:00000001`sur .

> [!IMPORTANT]
> Internet Explorer 11 ne prend pas en charge les versions de JavaScript ultérieures à la version ES5. Si l’un des utilisateurs de votre complément a des plateformes qui utilisent Internet Explorer 11, vous disposez de deux options pour utiliser la syntaxe et les fonctionnalités d’ECMAScript 2015 ou version ultérieure.
>
> - Écrivez votre code dans ECMAScript 2015 (également appelé ES6) ou javaScript ultérieur, ou en TypeScript, puis compilez votre code en JavaScript ES5 à l’aide d’un compilateur tel que [babel](https://babeljs.io/) ou [tsc](https://www.typescriptlang.org/index.html).
> - Écrivez dans ECMAScript 2015 ou version ultérieure de JavaScript, mais chargez également une bibliothèque [de polyfills](https://en.wikipedia.org/wiki/Polyfill_(programming)) telle que [core-js](https://github.com/zloirock/core-js) qui permet à Internet Explorer d’exécuter votre code.
>
> Pour plus d’informations sur ces options, consultez [Support Internet Explorer 11](../develop/support-ie-11.md).
>
> Par ailleurs, Internet Explorer 11 ne prend pas en charge certaines fonctionnalités HTML5 telles que les éléments multimédias, l’enregistrement et l’emplacement. Pour plus d’informations, consultez [Déterminer au moment de l’exécution si le complément est en cours d’exécution dans Internet Explorer](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

## <a name="troubleshooting-microsoft-edge-issues"></a>Résolution des problèmes de Microsoft Edge

### <a name="service-workers-are-not-working"></a>Les employés de service ne fonctionnent pas

Office compléments ne prennent pas en charge les Services Workers lorsque le Microsoft Edge WebView d’origine, [EdgeHTML](https://en.wikipedia.org/wiki/EdgeHTML), est utilisé. Ils sont pris en charge avec edge [WebView2 basé sur Chromium](/microsoft-edge/hosting/webview2).

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>La barre de défilement n’apparaît pas dans le volet des tâches

Par défaut, les barres de défilement dans Microsoft Edge sont masquées jusqu’au moment où elles sont survolées. Pour vous assurer que la barre de défilement est toujours visible, les styles CSS qui s’appliquent à l’`<body>`élément des pages dans le volet des tâches doivent inclure la propriété [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/Microsoft_Extensions) et la valeur `scrollbar` doit être attribuée.

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Lorsque vous déboguez avec Microsoft Edge DevTools, le complément se bloque ou se recharge

Le paramétrage de points d'arrêt dans [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) peut faire croire à Office que le complément est suspendu. Lorsque cela se produit, le complément est alors automatiquement rechargé. Pour éviter ce phénomène, ajoutez la valeur et la clé de registre suivantes à l’ordinateur de développement : `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>Lorsque le complément tente de s’ouvrir, l’erreur « ERREUR DE COMPLÉMENT Impossible d’ouvrir ce complément à partir de localhost » apparaît.

Microsoft Edge exige que localhost bénéficie d’une exemption de bouclage sur l’ordinateur de développement, ce qui est une raison connue. Suivez les instructions à l’emplacement suivant : [Impossible d’ouvrir le complément à partir de localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### <a name="get-errors-trying-to-download-a-pdf-file"></a>Obtenir des erreurs lors de la tentative de téléchargement d’un fichier PDF

Le téléchargement direct d’objets blob en tant que fichiers PDF dans un complément n’est pas pris en charge lorsque Edge est le navigateur. La solution de contournement consiste à créer une application web simple qui télécharge des objets blob en tant que fichiers PDF. Dans votre complément, appelez la `Office.context.ui.openBrowserWindow(url)` méthode et transmettez l’URL de l’application web. Cela ouvre l’application web dans une fenêtre de navigateur en dehors de Office.

## <a name="see-also"></a>Voir aussi

- [Configuration requise pour exécuter des compléments Office](requirements-for-running-office-add-ins.md)
