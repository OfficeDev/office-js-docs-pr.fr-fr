---
title: Navigateurs utilisés par les compléments Office
description: Indique comment le système d’exploitation et la version d’Office déterminent le navigateur utilisé par les compléments Office.
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: 0bd231cc870322dd6f756defd14e4d67a69478b4
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771248"
---
# <a name="browsers-used-by-office-add-ins"></a>Navigateurs utilisés par les compléments Office

Les compléments Office sont des applications Web qui sont affichées à l’aide d’iFrames lorsqu’ils sont exécutés dans Office sur le Web et qui utilisent des contrôles de navigateur incorporés dans Office pour les clients mobiles et de bureau. Les compléments ont également besoin d’un moteur JavaScript pour exécuter le code JavaScript. Le navigateur et le moteur incorporés sont fournis par un navigateur installé sur l’ordinateur de l’utilisateur.

Le navigateur utilisé dépend de ce qui suit :

- Système d’exploitation de l’ordinateur.
- Si le complément est exécuté dans Office sur le Web, Microsoft 365 ou sans abonnement Office 2013 ou version ultérieure.

Le tableau ci-dessous répertorie le navigateur utilisé selon les plateformes et systèmes d’exploitation.

|OS|Version d’Office|WebView2 Edge installé (basé sur le chrome) ?|Navigateur|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|indifférent|Office sur le web|Non applicable|Navigateur dans lequel Office sur le web est ouvert.|
|Mac|indifférent|Non applicable|Safari|
|iOS|indifférent|Non applicable|Safari|
|Android|indifférent|Non applicable|Chrome|
|Windows 7, 8,1, 10 | Office 2013 sans abonnement ou version ultérieure|Peu importe|Internet Explorer 11|
|Windows 7 | Microsoft 365| Peu importe | Internet Explorer 11|
|Windows 8,1,<br>Windows 10 ver. &nbsp; < &nbsp; 1903| Microsoft 365 | Non| Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; < &nbsp; 16.0.11629<sup>1</sup>| Peu importe|Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.11629 &nbsp; _et_ &nbsp; < &nbsp; 16.0.13530.20316 <sup>1</sup>| Peu importe|Microsoft Edge<sup>2, 3</sup> avec WebView d’origine (EdgeHTML)|
|Windows 10 ver. &nbsp; >= &nbsp; 1903 | Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13530.20316<sup>1</sup>| Non |Microsoft Edge<sup>2, 3</sup> avec WebView d’origine (EdgeHTML)|
|Windows 8.1<br>Windows 10| Microsoft 365 ver. &nbsp; >= &nbsp; 16.0.13530.20316<sup>1</sup>| Oui<sup>4</sup>|  Microsoft Edge<sup>2, 3</sup> avec WebView2 (basé sur le chrome) |

<sup>1</sup> pour plus d’informations, consultez la [page historique des mises à jour](/officeupdates/update-history-office365-proplus-by-date) et [Découvrez comment trouver votre version de client Office et votre canal de mise à jour](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19) .

<sup>2</sup> lorsque Microsoft Edge est utilisé, le narrateur Windows 10 (parfois appelé « lecteur d’écran ») lit la `<title>` balise dans la page qui s’ouvre dans le volet Office. Si Internet Explorer 11 est utilisé, le Narrateur lit la barre de titre du volet Office, qui provient de la valeur `<DisplayName>` du manifeste du complément.

<sup>3</sup> si votre complément inclut l' `Runtimes` élément dans le manifeste, il utilise Internet Explorer 11 quelle que soit la version de Windows ou de Microsoft 365. Pour plus d’informations, voir [Services d’exécution](../reference/manifest/runtimes.md).

<sup>4</sup> le contrôle WebView2 incorporable doit être installé en plus de l’installation de Microsoft Edge pour permettre à Office de l’incorporer. Pour l’installer, voir [Microsoft Edge WebView2/embed Web Content... avec Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/).


> [!IMPORTANT]
> Internet Explorer 11 ne prend pas en charge les versions de JavaScript ultérieures à la version ES5. Si les utilisateurs de votre complément disposent de plateformes qui utilisent Internet Explorer 11, pour utiliser la syntaxe et les fonctionnalités d’ECMAScript 2015 ou version ultérieure, vous disposez de deux options :
>
> - Écrivez votre code dans ECMAScript 2015 (également appelé ES6) ou JavaScript ultérieur, ou dans une écriture à écrire, puis compilez votre code en ES5 JavaScript à l’aide d’un compilateur tel que [Babel](https://babeljs.io/) ou [TSC](https://www.typescriptlang.org/index.html).
> - Écrivez dans ECMAScript 2015 ou une version ultérieure JavaScript, mais chargez également une bibliothèque de [Polyfill](https://wikipedia.org/wiki/Polyfill_(programming)) comme [Core-js](https://github.com/zloirock/core-js) qui permet à Internet Explorer d’exécuter votre code.
>
> Par ailleurs, Internet Explorer 11 ne prend pas en charge certaines fonctionnalités HTML5 telles que les éléments multimédias, l’enregistrement et l’emplacement.

## <a name="troubleshooting-microsoft-edge-issues"></a>Résolution des problèmes liés à Microsoft Edge

### <a name="service-workers-are-not-working"></a>Les travailleurs de services ne fonctionnent pas

Les compléments Office ne prennent pas en charge les travailleurs de services lorsque le WebView d’origine de [Microsoft Edge](/microsoft-edge/hosting/webview) est utilisé. Elles sont prises en charge avec le [WebView2 Edge basé sur le chrome](/microsoft-edge/hosting/webview2).

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>La barre de défilement n’apparaît pas dans le volet des tâches

Par défaut, les barres de défilement dans Microsoft Edge sont masquées jusqu’au moment où elles sont survolées. Pour vous assurer que la barre de défilement est toujours visible, les styles CSS qui s’appliquent à l’`<body>`élément des pages dans le volet des tâches doivent inclure la propriété [-ms-overflow-style](https://developer.mozilla.org/docs/Archive/Web/CSS/-ms-overflow-style) et la valeur `scrollbar` doit être attribuée.

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Lorsque vous déboguez avec Microsoft Edge DevTools, le complément se bloque ou se recharge

Le paramétrage de points d'arrêt dans [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) peut faire croire à Office que le complément est suspendu. Lorsque cela se produit, le complément est alors automatiquement rechargé. Pour éviter ce phénomène, ajoutez la valeur et la clé de registre suivantes à l’ordinateur de développement : `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>Lorsque le complément tente de s’ouvrir, l’erreur « ERREUR DE COMPLÉMENT Impossible d’ouvrir ce complément à partir de localhost » apparaît.

Microsoft Edge exige que localhost bénéficie d’une exemption de bouclage sur l’ordinateur de développement, ce qui est une raison connue. Suivez les instructions à l’emplacement suivant : [Impossible d’ouvrir le complément à partir de localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### <a name="get-errors-trying-to-download-a-pdf-file"></a>Obtenir des erreurs lors de la tentative de téléchargement d’un fichier PDF

Le téléchargement direct d’objets BLOB sous forme de fichiers PDF dans un complément n’est pas pris en charge lorsque le serveur Edge est le navigateur. La solution de contournement consiste à créer une application Web simple qui télécharge les objets BLOB en tant que fichiers PDF. Dans votre complément, appelez la `Office.context.ui.openBrowserWindow(url)` méthode et transmettez l’URL de l’application Web. Cette opération ouvre l’application Web dans une fenêtre de navigateur en dehors d’Office.

## <a name="see-also"></a>Voir aussi

- [Configuration requise pour exécuter des compléments Office](requirements-for-running-office-add-ins.md)
