---
title: Ensembles de conditions requises de l’API JavaScript pour Outlook
description: En savoir plus sur les ensembles de conditions requises de l’API JavaScript pour Outlook.
ms.date: 04/29/2021
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: 691c0ad3ba97d79103e06bc7c0832eb8327b817a
ms.sourcegitcommit: 6057afc1776e1667b231d2e9809d261d372151f6
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/30/2021
ms.locfileid: "52100291"
---
# <a name="outlook-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour Outlook

Les versions API requises pour les compléments Outlook sont indiquées à l’aide de l’élément Requirements dans leur manifeste. Les compléments Outlook contiennent toujours un élément Set avec un attribut  défini sur  et un attribut  défini sur l’ensemble minimal de conditions requises de l’API qui prend en charge les scénarios du complément.

Par exemple, l’extrait de manifeste suivant indique l’ensemble minimal de conditions requises 1.1.

```xml
<Requirements>
  <Sets>
    <Set Name="Mailbox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

Toutes les API Outlook appartiennent à l’[ensemble de conditions requises](../../develop/specify-office-hosts-and-api-requirements.md) `Mailbox`. L’ensemble de conditions requises `Mailbox` possède plusieurs versions et chaque nouvel ensemble d’API publié appartient à une version plus récente de l’ensemble. L’ensemble d’API le plus récent n’est pas pris en charge par tous les clients Outlook, mais si ce dernier prend en charge un ensemble de conditions requises, toutes les API comprises dans cet ensemble sont également prises en charge. (consultez la documentation sur une API ou une fonctionnalité spécifique pour les exceptions).

La définition d’une version minimale d’ensemble de conditions requises dans le manifeste permet de contrôler les clients Outlook dans lesquels le complément va apparaître. Si un client ne prend pas en charge l’ensemble minimal de conditions requises, il ne charge pas le complément. Par exemple, si la version de l’ensemble de conditions requises spécifiée est 1.3, le complément n’apparaîtra pas dans les clients Outlook qui ne prennent pas en charge au minimum la version 1.3.

> [!NOTE]
> Pour utiliser des API dans l’un des ensembles de conditions requises numérotés, vous devez référencer la bibliothèque de **production** sur le CDN (https://appsforoffice.microsoft.com/lib/1/hosted/office.js).
>
> Pour plus d’informations sur l’utilisation des API disponibles en préversion, consultez la section [Utilisation des API disponibles en préversion](#using-preview-apis) plus loin dans cet article.

## <a name="using-apis-from-later-requirement-sets"></a>Utilisation des API d’un ensemble de conditions requises ultérieure

La définition d’un ensemble de conditions requises ne limite pas votre complément à utiliser les API de cette version. Par exemple, si le complément indique l’ensemble de conditions « Boîte aux lettres 1.1 », mais qu’il s’est exécuté dans un client Outlook prenant en charge « Boîte aux lettres 1.3 », le complément peut utiliser les API de l’ensemble de conditions « Boîte aux lettres 1.3 ».

Pour utiliser une nouvelle API, les développeurs peuvent vérifier si une application particulière prend en charge l’ensemble des conditions requises en procédant comme suit :

```js
if (Office.context.requirements.isSetSupported('Mailbox', '1.3')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

Autrement, les développeurs peuvent vérifier la disponibilité d’une nouvelle API en utilisant la technique JavaScript standard.

```js
if (item.somePropertyOrFunction !== undefined) {
  // Use item.somePropertyOrFunction.
  item.somePropertyOrFunction;
}
```

Ces vérifications ne sont pas nécessaires pour les API présentes dans l’ensemble de conditions requises dont la version est la même que celle spécifiée dans le manifeste.

## <a name="choosing-a-minimum-requirement-set"></a>Choix d’un ensemble minimal de conditions requises

Les développeurs doivent utiliser l’ensemble de conditions requises le plus ancien qui contient l’ensemble d’API critique pour leur scénario, sans lequel le complément ne fonctionne pas.

## <a name="requirement-sets-supported-by-exchange-servers-and-outlook-clients"></a>Ensembles de conditions requises pris en charge par les serveurs Exchange et les clients Outlook

Dans cette section, nous prenons note de la plage d’ensembles de conditions requises pris en charge par les serveurs Exchange et les clients Outlook. Pour plus d’informations sur la configuration requise pour le serveur et le client pour l’exécution de compléments Outlook, voir [Conditions requises pour les compléments Outlook](../../outlook/add-in-requirements.md).

> [!IMPORTANT]
> Si votre serveur Exchange cible et votre client Outlook prennent en charge différents ensembles de conditions requises, vous êtes limité à la plage inférieure d’ensembles de conditions requises. Par exemple, si un complément est exécuté dans Outlook 2016 sur Mac (configuration maximale requise : 1.6) sur Exchange 2013 (ensemble de conditions requises le plus élevé : 1.1), votre complément est limité à l’ensemble de conditions requises 1.1.

### <a name="exchange-server-support"></a>Prise en charge par le serveur Exchange

Les serveurs suivants prennent en charge des compléments Outlook.

| Produit | Version principale d’Exchange | Ensembles de conditions requises des API prises en charge |
|---|---|---|
| Exchange Online | Dernière version | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](../objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)\* |
| Exchange local | 2019 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2016 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2013 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md) |

> [!NOTE]
> \* Pour nécessiter le jeu d'API d'identité 1.3 dans votre code additionnel, vérifiez s'il est pris en charge en appelant `isSetSupported('IdentityAPI', '1.3')`. Le déclarer dans le manifeste de votre macro complémentaire n'est pas pris en charge. Vous pouvez également déterminer si l’API est prise en charge en vérifiant qu’elle n’est pas `undefined`. Pour plus d’informations, consultez [Utilisation des API d’un ensemble de conditions requises ultérieure](#using-apis-from-later-requirement-sets).

### <a name="outlook-client-support"></a>Prise en charge du client Outlook

Les compléments sont pris en charge dans Outlook sur les plateformes suivantes.

| Plateforme | Version principale d’Office/Outlook | Ensembles de conditions requises des API prises en charge |
|---|---|---|
| Windows | Abonnement Microsoft 365 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<sup>1</sup>, [1.9](../objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)<sup>1</sup><br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| Achat définitif 2019 (au détail) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<sup>1</sup> |
|| Achat définitif 2019 (licence en volume) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md) |
|| Achat définitif 2016 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)<sup>3</sup> |
|| Achat définitif 2013 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)<sup>3</sup>, [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)<sup>3</sup> |
| Mac | interface utilisateur actuelle<br>(connecté à un abonnement Microsoft 365) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| new UI (preview)<sup>4</sup><br>(connecté à un abonnement Microsoft 365) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md) |
|| Achat définitif 2019 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
|| Achat définitif 2016 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
| iOS | Abonnement Microsoft 365 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)<sup>5</sup> |
| Android | Abonnement Microsoft 365 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)<sup>5</sup> |
| Navigateur web | interface utilisateur moderne d’Outlook lors de sa connexion à<br>Exchange Online : abonnement à Microsoft 365, Outlook.com | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](../objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| interface utilisateur classique d’Outlook lors de sa connexion à<br>Exchange local | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |

> [!NOTE]
> <sup>1</sup> La prise en charge de la version **1.8** dans Outlook sous Windows avec un abonnement Microsoft 365 ou un achat définitif au détail est disponible depuis la version 1910 (build 12130.20272). La prise en charge de la version **1.9** dans Outlook sous Windows avec un abonnement Microsoft 365 est disponible depuis la version 2008 (build 13127.20296). Pour plus d’informations en fonction de version, consultez la page de l’historique des mises à jour d’[Office 2019](/officeupdates/update-history-office-2019) ou [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) et [comment trouver la version client et le canal de mise à jour Office que vous utilisez](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).
>
> <sup>2</sup> Pour nécessiter le jeu d'API d'identité 1.3 dans votre code additionnel, vérifiez s'il est pris en charge en appelant `isSetSupported('IdentityAPI', '1.3')`. Le déclarer dans le manifeste de votre macro complémentaire n'est pas pris en charge. Vous pouvez également déterminer si l’API est prise en charge en vérifiant qu’elle n’est pas `undefined`. Pour plus d’informations, consultez [Utilisation des API d’un ensemble de conditions requises ultérieure](#using-apis-from-later-requirement-sets).
>
> <sup>3</sup> La prise en charge de la version 1.3 dans Outlook 2013 a été ajoutée dans le cadre de la [mise à jour du 8 décembre 2015 pour Outlook 2013 (KB3114349)](https://support.microsoft.com/kb/3114349). La prise en charge de la version 1.4 dans Outlook 2013 a été ajoutée dans le cadre de la [mise à jour du 13 septembre 2016 pour Outlook 2013 (KB3118280)](https://support.microsoft.com/help/3118280). La prise en charge de la version 1.4 dans Outlook 2016 a été ajoutée dans le cadre la [mise à jour du 3 juillet 2018 pour Office 2016 (KB4022223)](https://support.microsoft.com/help/4022223).
>
> <sup>4</sup> prise en charge de la nouvelle interface utilisateur (préversion) d’Outlook sur Mac est disponible dans la version 16.38.506. Pour plus d’informations, consultez la section [Prise en charge du macro complémentaire dans Outlook sur le nouvel interface d’utilisateur Mac](../../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview).
>
> <sup>5</sup> actuellement, d’autres aspects sont à prendre en considération lors de la conception et de l’implémentation de compléments pour les clients mobiles. Par exemple, le seul mode pris en charge est lecture de message. Pour plus d’informations, voir considérations relatives au code [lors de l’ajout de la prise en charge des commandes de complément pour les](../../outlook/add-mobile-support.md#code-considerations)Outlook Mobile.

> [!TIP]
> Vous pouvez faire la distinction entre les deux versions d’Outlook, classique et moderne, dans un navigateur Web en regardant la barre d’outils de votre boîte aux lettres.
>
> **moderne**
>
> ![capture d’écran partielle de la barre d’outils Outlook moderne](../../images/outlook-on-the-web-new-toolbar.png)
>
> **classique**
>
> ![capture d’écran partielle de la barre d’outils Outlook classique](../../images/outlook-on-the-web-classic-toolbar.png)

## <a name="using-preview-apis"></a>Utilisation des API de préversion

Les nouvelles API Outlook JavaScript sont d’abord introduites dans la « préversion », puis deviennent partie intégrante d’un ensemble de conditions requises spécifiques numérotées une fois qu’un nombre suffisant de tests a été effectué et que les utilisateurs ont renvoyé des commentaires. Pour formuler des commentaires sur une version d’évaluation API, utilisez le mécanisme de commentaires à la fin de la page web où l’API est documenté.

> [!NOTE]
> L’aperçu API peut être modifiés et n’est pas destinés à utiliser dans un environnement de production.

Pour plus d’informations sur les API de préversion, reportez-vous à l’article relatif à l’[ensemble de conditions requises de l’API Outlook de préversion](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md).
