---
title: Navigateurs utilisés par les compléments Office
description: Indique comment le système d’exploitation et la version d’Office déterminent le navigateur utilisé par les compléments Office.
ms.date: 07/27/2022
ms.localizationpriority: medium
ms.openlocfilehash: a4063720f8866d9538865f4514841d8dc8d0a84c
ms.sourcegitcommit: 143ab022c9ff6ba65bf20b34b5b3a5836d36744c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/03/2022
ms.locfileid: "67177671"
---
# <a name="browsers-used-by-office-add-ins"></a>Navigateurs utilisés par les compléments Office

Les compléments Office sont des applications web qui sont affichées à l’aide d’iFrames lors de l’exécution dans Office sur le Web. Dans Office pour les clients de bureau et mobiles, les compléments Office utilisent un contrôle de navigateur incorporé (également appelé webview). Les compléments ont également besoin d’un moteur JavaScript pour exécuter le code JavaScript. Le navigateur incorporé et le moteur sont fournis par un navigateur installé sur l’ordinateur de l’utilisateur.

Le navigateur utilisé dépend de ce qui suit :

- Système d’exploitation de l’ordinateur.
- Que le complément s’exécute dans Office sur le Web, Microsoft 365 ou perpétuel (également appelé « achat sans abonnement » ou « achat unique ») Office 2013 ou ultérieur.
- Dans les versions perpétuelles d’Office, que le complément s’exécute dans la variante « consommateur » ou « commercial » (également appelée « licence en volume » ou « LTSC »).

> [!IMPORTANT]
> **Internet Explorer toujours utilisé dans les compléments Office**
>
> Certaines combinaisons de plateformes et de versions d’Office, notamment les versions perpétuelles via Office 2019, utilisent toujours le contrôle webview fourni avec Internet Explorer 11 pour héberger des compléments, comme expliqué dans cet article. Nous vous recommandons (mais n’exigez pas) de continuer à prendre en charge ces combinaisons, du moins d’une manière minimale, en fournissant aux utilisateurs de votre complément un message d’échec approprié lorsque votre complément est lancé dans la vue web d’Internet Explorer. Gardez à l’esprit ces points supplémentaires :
>
> - Office sur le Web ne s’ouvre plus dans Internet Explorer. Par conséquent, [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) ne teste plus les compléments dans Office sur le Web à l’aide d’Internet Explorer comme navigateur.
> - AppSource teste toujours les combinaisons de versions de plateforme et de *bureau* Office qui utilisent Internet Explorer, mais elle émet uniquement un avertissement lorsque le complément ne prend pas en charge Internet Explorer; le complément n’est pas rejeté par AppSource.
> - [L’outil Script Lab](../overview/explore-with-script-lab.md) ne prend plus en charge Internet Explorer.
>
> Pour plus d’informations sur la prise en charge d’Internet Explorer et la configuration d’un message d’échec approprié sur votre complément, consultez [Support Internet Explorer 11](../develop/support-ie-11.md).

Les sections suivantes spécifient le navigateur utilisé pour les différentes plateformes et systèmes d’exploitation.

## <a name="non-windows-platforms"></a>Plateformes non Windows

Pour ces plateformes, la plateforme détermine à elle seule le navigateur utilisé.

|Système d’exploitation|Version d’Office|Navigateur|
|:-----|:-----|:-----|
|indifférent|Office sur le web|Navigateur dans lequel Office sur le web est ouvert.<br>(Notez toutefois que Office sur le Web ne s’ouvre pas dans Internet Explorer.<br>Si vous tentez de le faire, Office sur le Web s’ouvre dans Edge.) |
|Mac|indifférent|Safari avec WKWebView|
|iOS|indifférent|Safari avec WKWebView|
|Android|indifférent|Chrome|

## <a name="perpetual-versions-of-office-on-windows"></a>Versions perpétuelles d’Office sur Windows

Pour les versions perpétuelles d’Office sur Windows, le navigateur utilisé est déterminé par la version d’Office, si la licence est consommateur ou commerciale, et si edge WebView2 (basé sur Chromium) est installé. La version de Windows n’a pas d’importance, mais notez que les compléments Web Office ne sont pas pris en charge sur les versions antérieures à Windows 7 et Office 2021 n’est pas pris en charge sur les versions antérieures à Windows 10.

Pour déterminer si Office 2016 ou Office 2019 est consommateur ou commercial, utilisez le format de la version et du numéro de build d’Office. (Pour Office 2013 et Office 2021, la distinction entre commercial et consommateur n’a pas d’importance.)

- **Consommateur** : pour Office 2016 et 2019, le format est `YYMM (xxxxx.xxxxxx)`, se terminant par deux blocs de cinq chiffres ; par exemple, `2206 (Build 15330.20264`.
- **Commercial** : 

    - Pour Office 2016, le format est `16.0.xxxx.xxxxx`, se terminant par deux blocs de *quatre* chiffres ; par exemple, `16.0.5197.1000`.
    - Pour Office 2019, le format est `1808 (xxxxx.xxxxxx)`, se terminant par deux blocs de *cinq* chiffres ; par exemple, `1808 (Build 10388.20027)`. Notez que l’année et le mois sont toujours `1808`.

|Version d’Office|Consommateur et commercial|Edge WebView2 (basé sur Chromium) installé ?|Navigateur|
|:-----|:-----|:-----|:-----|:-----|
|Office 2013 | Peu importe |Peu importe|Internet Explorer 11|
|Office 2016| Commerciale |Peu importe|Internet Explorer 11|
|Office 2019| Commerciale |Peu importe|Internet Explorer 11|
|Office 2016 vers Office 2019| Grand public |Non |Microsoft Edge<sup>1, 2</sup> avec WebView d’origine (EdgeHTML)</br>Si Edge n’est pas installé, Internet Explorer 11 est utilisé.|
|Office 2016 vers Office 2019|  Grand public |Oui<sup>3</sup>|Microsoft Edge<sup>1</sup> avec WebView2 (basé sur Chromium)|
|Office 2021| Peu importe |Oui<sup>3</sup> |Microsoft Edge<sup>1</sup> avec WebView2 (basé sur Chromium)|

<sup>1</sup> Lorsque vous utilisez Microsoft Edge, le Narrateur Windows (parfois appelé « lecteur d’écran ») lit la `<title>` balise dans la page qui s’ouvre dans le volet Office. Dans Internet Explorer 11, le Narrateur lit la barre de titre du volet Office, qui provient de la **\<DisplayName\>** valeur dans le manifeste du complément.

<sup>2</sup> Si votre complément inclut l’élément **\<Runtimes\>** dans le manifeste, il n’utilisera pas Microsoft Edge avec le WebView d’origine (EdgeHTML). Si les conditions d’utilisation de Microsoft Edge avec WebView2 (basée sur Chromium) sont remplies, le complément utilise ce navigateur. Sinon, il utilise Internet Explorer 11. Pour plus d’informations, voir [Services d’exécution](/javascript/api/manifest/runtimes).

<sup>3</sup> Sur les versions windows antérieures à Windows 11, le contrôle WebView2 doit être installé afin qu’Office puisse l’incorporer. Il est installé avec des Office 2021 perpétuelles ou ultérieures, mais il n’est pas installé automatiquement avec Microsoft Edge. Si vous disposez d’une version antérieure d’Office perpétuel, suivez les instructions d’installation du contrôle sur [Microsoft Edge WebView2 / Incorporer du contenu web ... avec Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/).

## <a name="microsoft-365-subscription-on-windows"></a>Abonnement Microsoft 365 sur Windows

Pour l’abonnement Office sur Windows, le navigateur utilisé est déterminé par le système d’exploitation, la version d’Office et si edge WebView2 (basé sur Chromium) est installé.

|Système d’exploitation|Version d’Office|Edge WebView2 (basé sur Chromium) installé ?|Navigateur|
|:-----|:-----|:-----|:-----|
|Windows 7 | Microsoft 365| Peu importe | Internet Explorer 11|
|Windows 8.1,<br>Windows 10 ver.&nbsp;<&nbsp; 1903| Microsoft 365 | Non| Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp; 1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;<&nbsp; 16.0.11629<sup>2</sup>| Peu importe|Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp; 1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;>=&nbsp; 16.0.11629&nbsp;_ET_&nbsp;<&nbsp;16.0.13530.20424 <sup>2</sup>| Peu importe|Microsoft Edge<sup>1, 3</sup> avec WebView d’origine (EdgeHTML)|
|Windows 10 ver.&nbsp;>=&nbsp; 1903,<br>Fenêtre 11 | Microsoft 365 ver.&nbsp;>=&nbsp; 16.0.13530.20424<sup>2</sup>| Non |Microsoft Edge<sup>1, 3</sup> avec WebView d’origine (EdgeHTML)|
|Windows 8.1<br>Windows 10,<br>Windows 11| Microsoft 365 ver.&nbsp;>=&nbsp; 16.0.13530.20424<sup>2</sup>| Oui<sup>4</sup>|  Microsoft Edge<sup>1</sup> avec WebView2 (basé sur Chromium) |

<sup>1</sup> Lorsque vous utilisez Microsoft Edge, le Narrateur Windows (parfois appelé « lecteur d’écran ») lit la `<title>` balise dans la page qui s’ouvre dans le volet Office. Dans Internet Explorer 11, le Narrateur lit la barre de titre du volet Office, qui provient de la **\<DisplayName\>** valeur dans le manifeste du complément.

<sup>2</sup> Pour plus d’informations, consultez la page historique des [mises à jour](/officeupdates/update-history-office365-proplus-by-date) et comment [trouver la version et le canal de mise à jour de votre client Office](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19) .

<sup>3</sup> Si votre complément inclut l’élément **\<Runtimes\>** dans le manifeste, il n’utilisera pas Microsoft Edge avec le WebView d’origine (EdgeHTML). Si les conditions d’utilisation de Microsoft Edge avec WebView2 (basée sur Chromium) sont remplies, le complément utilise ce navigateur. Sinon, il utilise Internet Explorer 11, quelle que soit la version de Windows ou Microsoft 365. Pour plus d’informations, voir [Services d’exécution](/javascript/api/manifest/runtimes).

<sup>4</sup> Sur les versions windows antérieures à Windows 11, le contrôle WebView2 doit être installé afin qu’Office puisse l’incorporer. Il est installé avec Microsoft 365, version 2101 ou ultérieure, mais il n’est pas installé automatiquement avec Microsoft Edge. Si vous disposez d’une version antérieure de Microsoft 365, suivez les instructions d’installation du contrôle sur [Microsoft Edge WebView2 / Incorporer du contenu web ... avec Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/). Sur les builds Microsoft 365 antérieures à 16.0.14326.xxxxx, vous devez également créer la clé de Registre **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Win32WebView2** et définir sa valeur `dword:00000001`sur .

## <a name="working-with-internet-explorer"></a>Utilisation d’Internet Explorer

Internet Explorer 11 ne prend pas en charge les versions de JavaScript ultérieures à la version ES5. Si l’un des utilisateurs de votre complément a des plateformes qui utilisent Internet Explorer 11, vous disposez de deux options pour utiliser la syntaxe et les fonctionnalités d’ECMAScript 2015 ou version ultérieure.

- Écrivez votre code dans ECMAScript 2015 (également appelé ES6) ou javaScript ultérieur, ou en TypeScript, puis compilez votre code en JavaScript ES5 à l’aide d’un compilateur tel que [babel](https://babeljs.io/) ou [tsc](https://www.typescriptlang.org/index.html).
- Écrivez dans ECMAScript 2015 ou version ultérieure de JavaScript, mais chargez également une bibliothèque [de polyfills](https://en.wikipedia.org/wiki/Polyfill_(programming)) telle que [core-js](https://github.com/zloirock/core-js) qui permet à Internet Explorer d’exécuter votre code.

Pour plus d’informations sur ces options, consultez [Support Internet Explorer 11](../develop/support-ie-11.md).

Par ailleurs, Internet Explorer 11 ne prend pas en charge certaines fonctionnalités HTML5 telles que les éléments multimédias, l’enregistrement et l’emplacement. Pour plus d’informations, consultez [Déterminer au moment de l’exécution si le complément est en cours d’exécution dans Internet Explorer](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

## <a name="troubleshoot-microsoft-edge-issues"></a>Résoudre les problèmes liés à Microsoft Edge

### <a name="service-workers-are-not-working"></a>Les employés de service ne fonctionnent pas

Les compléments Office ne prennent pas en charge les Employés de service lorsque le WebView Microsoft Edge d’origine, [EdgeHTML](https://en.wikipedia.org/wiki/EdgeHTML), est utilisé. Ils sont pris en charge avec edge [WebView2 basé sur Chromium](/microsoft-edge/hosting/webview2).

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>La barre de défilement n’apparaît pas dans le volet des tâches

Par défaut, les barres de défilement dans Microsoft Edge sont masquées jusqu’au moment où elles sont survolées. Pour vous assurer que la barre de défilement est toujours visible, les styles CSS qui s’appliquent à l’`<body>`élément des pages dans le volet des tâches doivent inclure la propriété [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/Microsoft_Extensions) et la valeur `scrollbar` doit être attribuée.

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Lorsque vous déboguez avec Microsoft Edge DevTools, le complément se bloque ou se recharge

Le paramétrage de points d'arrêt dans [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) peut faire croire à Office que le complément est suspendu. Lorsque cela se produit, le complément est alors automatiquement rechargé. Pour éviter ce phénomène, ajoutez la valeur et la clé de registre suivantes à l’ordinateur de développement : `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>Lorsque le complément tente de s’ouvrir, l’erreur « ERREUR DE COMPLÉMENT Impossible d’ouvrir ce complément à partir de localhost » apparaît.

Microsoft Edge exige que localhost bénéficie d’une exemption de bouclage sur l’ordinateur de développement, ce qui est une raison connue. Suivez les instructions à l’emplacement suivant : [Impossible d’ouvrir le complément à partir de localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### <a name="get-errors-trying-to-download-a-pdf-file"></a>Obtenir des erreurs lors de la tentative de téléchargement d’un fichier PDF

Le téléchargement direct d’objets blob en tant que fichiers PDF dans un complément n’est pas pris en charge lorsque Edge est le navigateur. La solution de contournement consiste à créer une application web simple qui télécharge des objets blob en tant que fichiers PDF. Dans votre complément, appelez la `Office.context.ui.openBrowserWindow(url)` méthode et transmettez l’URL de l’application web. Cela ouvre l’application web dans une fenêtre de navigateur en dehors d’Office.

## <a name="see-also"></a>Voir aussi

- [Configuration requise pour exécuter des compléments Office](requirements-for-running-office-add-ins.md)
