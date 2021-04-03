---
title: Navigateurs utilisés par les compléments Office
description: Indique comment le système d’exploitation et la version d’Office déterminent le navigateur utilisé par les compléments Office.
ms.date: 03/24/2021
localization_priority: Normal
ms.openlocfilehash: b9f4d07122779a893bd10e8d28b4f1b329125630
ms.sourcegitcommit: 074526a6dca8381dbdabf2705474c5ae6753b829
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51506132"
---
# <a name="browsers-used-by-office-add-ins"></a>Navigateurs utilisés par les compléments Office

Les add-ins Office sont des applications web qui s’affichent à l’aide d’iFrames lors de l’exécution dans Office sur le web et à l’aide de contrôles de navigateur incorporés dans Office pour les clients de bureau et mobiles. Les compléments ont également besoin d’un moteur JavaScript pour exécuter le code JavaScript. Le navigateur incorporé et le moteur sont fournis par un navigateur installé sur l’ordinateur de l’utilisateur.

Le navigateur utilisé dépend de ce qui suit :

- Système d’exploitation de l’ordinateur.
- Si le module est en cours d’exécution dans Office sur le web, Microsoft 365 ou Office 2013 sans abonnement ou ultérieur.

Le tableau ci-dessous répertorie le navigateur utilisé selon les plateformes et systèmes d’exploitation.

|SYSTÈME D’EXPLOITATION|Version d’Office|Edge WebView2 (basé sur Chromium) installé ?|Navigateur|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|indifférent|Office sur le web|Non applicable|Navigateur dans lequel Office sur le web est ouvert.|
|Mac|indifférent|Non applicable|Safari|
|iOS|indifférent|Non applicable|Safari|
|Android|indifférent|Non applicable|Chrome|
|Windows 7, 8.1, 10 | Office 2013 sans abonnement ou ultérieur|Peu importe|Internet Explorer 11|
|Windows 7 | Microsoft 365| Peu importe | Internet Explorer 11|
|Windows 8.1,<br>Windows 10 ver. &nbsp; < &nbsp; 1903| Microsoft 365 | Non| Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; < &nbsp; 16.0.11629<sup>1</sup>| Peu importe|Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.11629 &nbsp; _ET_ &nbsp; < &nbsp; 16.0.13530.20424 <sup>1</sup>| Peu importe|Microsoft Edge<sup>2, 3 avec</sup> WebView d’origine (EdgeHTML)|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13530.20424<sup>1</sup>| Non |Microsoft Edge<sup>2, 3 avec</sup> WebView d’origine (EdgeHTML)|
|Windows 8.1<br>Windows 10| Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13530.20424<sup>1</sup>| Oui<sup>4</sup>|  Microsoft Edge<sup>2</sup> avec WebView2 (basé sur Chromium) |

<sup>1 Consultez</sup> la [page historique des](/officeupdates/update-history-office365-proplus-by-date) mises à jour et découvrez comment trouver la version de votre client [Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19) et le canal de mise à jour pour plus d’informations.

<sup>2</sup> Lorsque Microsoft Edge est utilisé, le Narrateur Windows 10 (parfois appelé « lecteur d’écran ») lit la balise dans la page qui s’ouvre dans le volet `<title>` Des tâches. Si Internet Explorer 11 est utilisé, le Narrateur lit la barre de titre du volet Office, qui provient de la valeur `<DisplayName>` du manifeste du complément.

<sup>3</sup> Si votre application inclut l’élément dans le manifeste, elle n’utilisera pas Microsoft Edge avec le `<Runtimes>` WebView d’origine (EdgeHTML). Si les conditions d’utilisation de Microsoft Edge avec WebView2 (basé sur Chromium) sont remplies, le add-in utilise ce navigateur. Dans le cas contraire, il utilise Internet Explorer 11, quelle que soit la version de Windows ou de Microsoft 365. Pour plus d’informations, voir [Services d’exécution](../reference/manifest/runtimes.md).

<sup>4</sup> Le contrôle WebView2 intégrable doit être installé en plus de l’installation de Microsoft Edge afin qu’Office puisse l’incorporer. Pour l’installer, [voir Microsoft Edge WebView2 / Incorporer du contenu web... avec Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/).




> [!IMPORTANT]
> Internet Explorer 11 ne prend pas en charge les versions de JavaScript ultérieures à la version ES5. Si l’un des utilisateurs de votre add-in dispose de plateformes qui utilisent Internet Explorer 11, vous disposez de deux options pour utiliser la syntaxe et les fonctionnalités d’ECMAScript 2015 ou une ultérieure :
>
> - Écrivez votre code dans ECMAScript 2015 (également appelé ES6) ou version ultérieure JavaScript, ou dans TypeScript, puis compilez votre code en JavaScript ES5 à l’aide d’un compilateur tel que [celui-ci ou](https://babeljs.io/) [tsc.](https://www.typescriptlang.org/index.html)
> - Écrivez en JavaScript ECMAScript 2015 ou version ultérieure, mais chargez également une [bibliothèque polyfill](https://wikipedia.org/wiki/Polyfill_(programming)) telle que [core-js](https://github.com/zloirock/core-js) qui permet à IE d’exécuter votre code.
>
> Par ailleurs, Internet Explorer 11 ne prend pas en charge certaines fonctionnalités HTML5 telles que les éléments multimédias, l’enregistrement et l’emplacement.

## <a name="troubleshooting-microsoft-edge-issues"></a>Résolution des problèmes de Microsoft Edge

### <a name="service-workers-are-not-working"></a>Les employés de service ne fonctionnent pas

Les add-ins Office ne sont pas en charge par les travailleurs de service lorsque le [Microsoft Edge WebView](/microsoft-edge/hosting/webview) d’origine est utilisé. Ils sont pris en charge avec le [edge WebView2 basé sur Chromium.](/microsoft-edge/hosting/webview2)

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>La barre de défilement n’apparaît pas dans le volet des tâches

Par défaut, les barres de défilement dans Microsoft Edge sont masquées jusqu’au moment où elles sont survolées. Pour vous assurer que la barre de défilement est toujours visible, les styles CSS qui s’appliquent à l’`<body>`élément des pages dans le volet des tâches doivent inclure la propriété [-ms-overflow-style](https://developer.mozilla.org/docs/Archive/Web/CSS/-ms-overflow-style) et la valeur `scrollbar` doit être attribuée.

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Lorsque vous déboguez avec Microsoft Edge DevTools, le complément se bloque ou se recharge

Le paramétrage de points d'arrêt dans [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) peut faire croire à Office que le complément est suspendu. Lorsque cela se produit, le complément est alors automatiquement rechargé. Pour éviter ce phénomène, ajoutez la valeur et la clé de registre suivantes à l’ordinateur de développement : `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>Lorsque le complément tente de s’ouvrir, l’erreur « ERREUR DE COMPLÉMENT Impossible d’ouvrir ce complément à partir de localhost » apparaît.

Microsoft Edge exige que localhost bénéficie d’une exemption de bouclage sur l’ordinateur de développement, ce qui est une raison connue. Suivez les instructions à l’emplacement suivant : [Impossible d’ouvrir le complément à partir de localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### <a name="get-errors-trying-to-download-a-pdf-file"></a>Obtenir des erreurs lors de la tentative de téléchargement d’un fichier PDF

Le téléchargement direct des blobs en tant que fichiers PDF dans un add-in n’est pas pris en charge lorsque Edge est le navigateur. La solution de contournement consiste à créer une application web simple qui télécharge les blobs sous forme de fichiers PDF. Dans votre application, appelez la `Office.context.ui.openBrowserWindow(url)` méthode et passez l’URL de l’application web. Cette procédure ouvre l’application web dans une fenêtre de navigateur en dehors d’Office.

## <a name="see-also"></a>Voir aussi

- [Configuration requise pour exécuter des compléments Office](requirements-for-running-office-add-ins.md)
