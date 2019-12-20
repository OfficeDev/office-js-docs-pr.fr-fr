---
title: Navigateurs utilisés par les compléments Office
description: Indique comment le système d’exploitation et la version d’Office déterminent le navigateur utilisé par les compléments Office.
ms.date: 12/13/2019
localization_priority: Priority
ms.openlocfilehash: 3709157449634dfb49805e2247e47debe60f468f
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40813983"
---
# <a name="browsers-used-by-office-add-ins"></a>Navigateurs utilisés par les compléments Office

Les compléments Office sont des applications web qui s’affichent à l’aide d’iFrames lorsqu’ils sont exécutés dans Office sur le web et utilisent des contrôles de navigateur incorporés dans Office pour les clients de bureau et mobiles. Les compléments ont également besoin d’un moteur JavaScript pour exécuter le code JavaScript. Le navigateur incorporé et le moteur sont fournis par un navigateur installé sur l’ordinateur de l’utilisateur.

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
|Windows 10 version >= 1903/Office 365 version < 16.0.11629|Internet Explorer 11|
|Windows 10 version >= 1903/Office 365 version >= 16.0.11629|Microsoft Edge\*|

\*Si Microsoft Edge est utilisé, le Narrateur Windows 10 (parfois appelé « lecteur d’écran ») lit la balise `<title>` de la page qui s’ouvre dans le volet Office. Si Internet Explorer 11 est utilisé, le Narrateur lit la barre de titre du volet Office, qui provient de la valeur `<DisplayName>` du manifeste du complément.

> [!IMPORTANT]
> Internet Explorer 11 ne prend pas en charge les versions de JavaScript ultérieures à la version ES5. Si un des utilisateurs de votre complément dispose d’une plateforme utilisant Internet Explorer 11, vous devez transpiler JavaScript vers la version ES5 ou utiliser un polyfill pour lui permettre d’utiliser la syntaxe et les fonctionnalités d’ECMAScript 2015 ou version ultérieure. Par ailleurs, Internet Explorer 11 ne prend pas en charge certaines fonctionnalités HTML5 telles que les éléments multimédias, l’enregistrement et l’emplacement.

## <a name="troubleshooting-microsoft-edge-issues"></a>Résolution des problèmes liés à Microsoft Edge

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>La barre de défilement n’apparaît pas dans le volet des tâches

Par défaut, les barres de défilement dans Microsoft Edge sont masquées jusqu’au moment où elles sont survolées. Pour vous assurer que la barre de défilement est toujours visible, les styles CSS qui s’appliquent à l’`<body>`élément des pages dans le volet des tâches doivent inclure la propriété [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) et la valeur `scrollbar` doit être attribuée. 

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Lorsque vous déboguez avec Microsoft Edge DevTools, le complément se bloque ou se recharge

Le paramétrage de points d'arrêt dans [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) peut faire croire à Office que le complément est suspendu. Lorsque cela se produit, le complément est alors automatiquement rechargé. Pour éviter ce phénomène, ajoutez la valeur et la clé de registre suivantes à l’ordinateur de développement : `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>Lorsque le complément tente de s’ouvrir, l’erreur « ERREUR DE COMPLÉMENT Impossible d’ouvrir ce complément à partir de localhost » apparaît.

Microsoft Edge exige que localhost bénéficie d’une exemption de bouclage sur l’ordinateur de développement, ce qui est une raison connue. Suivez les instructions à l’emplacement suivant : [Impossible d’ouvrir le complément à partir de localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).


## <a name="see-also"></a>Voir aussi

- [Configuration requise pour exécuter des compléments Office](requirements-for-running-office-add-ins.md)
