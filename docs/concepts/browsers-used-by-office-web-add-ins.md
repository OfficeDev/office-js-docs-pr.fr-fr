---
title: Navigateurs utilisés par les compléments Office
description: Indique comment le système d’exploitation et la version d’Office déterminent le navigateur utilisé par les compléments Office.
ms.date: 04/21/2020
localization_priority: Normal
ms.openlocfilehash: 9ef4b6d4c09140fc6d6bb04eca51d845b79b6dc7
ms.sourcegitcommit: 3355c6bd64ecb45cea4c0d319053397f11bc9834
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/22/2020
ms.locfileid: "43744851"
---
# <a name="browsers-used-by-office-add-ins"></a>Navigateurs utilisés par les compléments Office

Les compléments Office sont des applications web qui s’affichent à l’aide d’iFrames lorsqu’ils sont exécutés dans Office sur le web et utilisent des contrôles de navigateur incorporés dans Office pour les clients de bureau et mobiles. Les compléments ont également besoin d’un moteur JavaScript pour exécuter le code JavaScript. Le navigateur et le moteur incorporés sont fournis par un navigateur installé sur l’ordinateur de l’utilisateur.

Le navigateur utilisé dépend de ce qui suit :

- Système d’exploitation de l’ordinateur.
- Exécution du complément dans Office sur le web, Office 365, Office 2013 sans abonnement ou version ultérieure.

Le tableau ci-dessous répertorie le navigateur utilisé selon les plateformes et systèmes d’exploitation.

|**Système d’exploitation/Plateforme**|**Navigateur**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|Office sur le web|Navigateur dans lequel Office sur le web est ouvert.|
|Mac|Safari|
|iOS|Safari|
|Android|Chrome|
|Windows/Office 2013 sans abonnement ou version ultérieure|Internet Explorer 11|
|Windows 10 version < 1903/Office 365|Internet Explorer 11|
|Windows 10 version >= 1903/Office 365 ver < 16.0.11629<sup>1</sup>|Internet Explorer 11|
|Windows 10 version >= 1903/Office 365 ver >= 16.0.11629<sup>1</sup>|Microsoft Edge<sup>2</sup>|

<sup>1</sup> pour plus d’informations, consultez la [page historique des mises à jour](/officeupdates/update-history-office365-proplus-by-date) et [Découvrez comment trouver votre version de client Office et votre canal de mise à jour](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19) .

<sup>2</sup> lorsque Microsoft Edge est utilisé, le narrateur Windows 10 (parfois appelé « lecteur d’écran ») lit la `<title>` balise dans la page qui s’ouvre dans le volet Office. Si Internet Explorer 11 est utilisé, le Narrateur lit la barre de titre du volet Office, qui provient de la valeur `<DisplayName>` du manifeste du complément.

> [!IMPORTANT]
> Internet Explorer 11 ne prend pas en charge les versions de JavaScript ultérieures à la version ES5. Si un des utilisateurs de votre complément dispose d’une plateforme utilisant Internet Explorer 11, vous devez transpiler JavaScript vers la version ES5 ou utiliser un polyfill pour lui permettre d’utiliser la syntaxe et les fonctionnalités d’ECMAScript 2015 ou version ultérieure. Par ailleurs, Internet Explorer 11 ne prend pas en charge certaines fonctionnalités HTML5 telles que les éléments multimédias, l’enregistrement et l’emplacement.

## <a name="troubleshooting-microsoft-edge-issues"></a>Résolution des problèmes liés à Microsoft Edge

### <a name="service-workers-are-not-working"></a>Les travailleurs de services ne fonctionnent pas

Les compléments Office ne prennent pas en charge les travailleurs de service sur [Microsoft Edge WebView](/microsoft-edge/hosting/webview). Consultez la rubrique [vue d’ensemble des compléments Office](../overview/office-add-ins.md) pour les dernières fonctionnalités prises en charge sur le contrôle Edge WebView. Nous travaillons difficilement à mettre en place la nouvelle [WebView2 Edge basée](/microsoft-edge/hosting/webview2) sur le chrome à la plateforme de compléments Office, dont nous pensons qu’elle prendra en charge les travailleurs de service.

### <a name="chromium-based-edge-is-installed-on-my-development-computer-but-my-add-in-does-not-use-it"></a>Le serveur Edge basé sur le chrome est installé sur mon ordinateur de développement, mais mon complément ne l’utilise pas

Le navigateur de base dans [Microsoft Edge](https://support.microsoft.com/help/4501095/download-the-new-microsoft-edge-based-on-chromium) est passé à chrome. L’ancienne base, appelée EdgeHTML, n’est pas supprimée lorsque le serveur Edge basé sur le chrome est installé. Office continuera à utiliser la base EdgeHTML pour les compléments jusqu’à ce qu’une version d’Office 365 qui prenne en charge le chrome soit installée sur l’ordinateur. Nous prévoyons que ces builds doivent être expédiées dans 2020. Elles apparaîtront probablement dans le canal Insiders dans le premier semestre.

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>La barre de défilement n’apparaît pas dans le volet des tâches

Par défaut, les barres de défilement dans Microsoft Edge sont masquées jusqu’au moment où elles sont survolées. Pour vous assurer que la barre de défilement est toujours visible, les styles CSS qui s’appliquent à l’`<body>`élément des pages dans le volet des tâches doivent inclure la propriété [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) et la valeur `scrollbar` doit être attribuée. 

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Lorsque vous déboguez avec Microsoft Edge DevTools, le complément se bloque ou se recharge

Le paramétrage de points d'arrêt dans [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) peut faire croire à Office que le complément est suspendu. Lorsque cela se produit, le complément est alors automatiquement rechargé. Pour éviter ce phénomène, ajoutez la valeur et la clé de registre suivantes à l’ordinateur de développement : `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>Lorsque le complément tente de s’ouvrir, l’erreur « ERREUR DE COMPLÉMENT Impossible d’ouvrir ce complément à partir de localhost » apparaît.

Microsoft Edge exige que localhost bénéficie d’une exemption de bouclage sur l’ordinateur de développement, ce qui est une raison connue. Suivez les instructions à l’emplacement suivant : [Impossible d’ouvrir le complément à partir de localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).

### <a name="get-errors-trying-to-download-a-pdf-file"></a>Obtenir des erreurs lors de la tentative de téléchargement d’un fichier PDF

Le téléchargement direct d’objets BLOB sous forme de fichiers PDF dans un complément n’est pas pris en charge lorsque le serveur Edge est le navigateur. La solution de contournement consiste à créer une application Web simple qui télécharge les objets BLOB en tant que fichiers PDF. Dans votre complément, appelez la `Office.context.ui.openBrowserWindow(url)` méthode et transmettez l’URL de l’application Web. Cette opération ouvre l’application Web dans une fenêtre de navigateur en dehors d’Office.

## <a name="see-also"></a>Voir aussi

- [Configuration requise pour exécuter des compléments Office](requirements-for-running-office-add-ins.md)
