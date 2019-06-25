---
title: Navigateurs utilisés par les compléments Office
description: Indique comment le système d’exploitation et la version d’Office déterminent le navigateur utilisé par les compléments Office.
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 56b74c0e43c8e9709ecd03a8c60a89d3869e44f8
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128107"
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

> [!NOTE]
> En attendant leur mise à la disposition générale, vous devez participer au programme Windows Insider pour obtenir Windows 1903 ou version ultérieure, ainsi qu’au programme Office Insider pour obtenir la version 16.0.11629 ou ultérieure.
>
> Pour participer au programme Windows Insider :
> 
> 1. Accédez à [Windows Insider](https://insider.windows.com) et cliquez sur le lien pour participer au programme Windows Insider.
> 2. Vous accédez alors à une page d’instructions sur l’utilisation des paramètres Windows pour activer les builds Windows. Suivez les instructions. Lorsque vous sélectionnez le rythme des mises à jour, choisissez l’option la plus rapide.
>
> Pour participer au programme Office Insider :
> 
> 1. Accédez à [Participer au programme Office Insider](https://insider.office.com/join).
> 2. Suivez les instructions détaillées sur cette page. Lorsque vous êtes invité à spécifier un canal, sélectionnez Insider.

## <a name="see-also"></a>Voir aussi

- [Configuration requise pour exécuter des compléments Office](requirements-for-running-office-add-ins.md)
