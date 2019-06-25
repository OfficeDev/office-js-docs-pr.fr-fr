---
title: Ensembles de conditions requises de l’API JavaScript pour Outlook
description: ''
ms.date: 06/20/2019
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: ffd6cb33c0b3c21d769b8551d798bed3ab3390fb
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127008"
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

Toutes les API Outlook appartiennent à l’`Mailbox`[ensemble de conditions requises](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements). L’ensemble de conditions requises `Mailbox` possède plusieurs versions et chaque nouvel ensemble d’API publié appartient à une version supérieure de l’ensemble. L’ensemble d’API le plus récent n’est pas pris en charge par tous les clients Outlook, mais si un client Outlook prend en charge un ensemble de conditions requises, toutes les API comprises dans cet ensemble sont également prises en charge.

La définition d’une version minimale d’ensemble de conditions requises dans le manifeste permet de contrôler les clients Outlook dans lesquels le complément va apparaître. Si un client ne prend pas en charge l’ensemble minimal de conditions requises, il ne charge pas le complément. Par exemple, si la version de l’ensemble de conditions requises spécifiée est 1.3, le complément n’apparaîtra pas dans les clients Outlook qui ne prennent pas en charge au minimum la version 1.3.

## <a name="using-apis-from-later-requirement-sets"></a>Utilisation des API d’un ensemble de conditions requises ultérieure

La définition d’un ensemble de conditions requises ne limite pas votre complément à utiliser les API de cette version. Par exemple, si le complément indique l’ensemble de conditions requises 1.1, mais qu’il s’est exécuté dans un client Outlook prenant en charge la version 1.3, le complément peut utiliser les API de l’ensemble de conditions requises 1.3.

Pour utiliser une nouvelle API, les développeurs peuvent vérifier si un hôte particulier prend en charge un ensemble de conditions requises en procédant comme suit.

```js
if (Office.context.requirements.isSetSupported('Mailbox', 1.3) === true) {
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

## <a name="clients"></a>Clients

Les clients suivants prennent en charge des compléments Outlook.

| Client | Ensembles de conditions requises des API prises en charge |
| --- | --- |
| Outlook sur Windows (connecté à l’abonnement Office 365) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2019 pour Windows (achat unique) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2016 pour Windows (achat unique) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook 2013 pour Windows (achat unique) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook sur Mac (connecté à l’abonnement Office 365) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2019 sur Mac (achat unique) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook 2016 sur Mac (achat unique) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook sur iOS | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook sur Android | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook sur le web (nouveau) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook sur le web (classique) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| N’importe quel client Outlook connecté à Exchange 2019 en local | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| N’importe quel client Outlook connecté à Exchange 2016 en local | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3) |
| N’importe quel client Outlook connecté à Exchange 2013 en local | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1) |

> [!NOTE]
> La prise en charge de la version 1.3 dans Outlook 2013 a été ajoutée dans le cadre de la [mise à jour du 8 décembre 2015 pour Outlook 2013 (KB3114349)](https://support.microsoft.com/kb/3114349). La prise en charge de la version 1.4 dans Outlook 2013 a été ajoutée dans le cadre de la [mise à jour du 13 septembre 2016 pour Outlook 2013 (KB3118280)](https://support.microsoft.com/help/3118280). La prise en charge 1.4 dans Outlook 2016 (MSI) a été ajouté dans le cadre de la [mise à jour du 3 juillet 2018, pour Office 2016 (KB4022223)](https://support.microsoft.com/help/4022223).

## <a name="using-preview-apis"></a>Utilisation des API de préversion

Les nouvelles API Outlook JavaScript sont d’abord introduites dans la « préversion », puis deviennent partie intégrante d’un ensemble de conditions requises spécifiques numérotées une fois qu’un nombre suffisant de tests a été effectué et que les utilisateurs ont renvoyé des commentaires. Pour formuler des commentaires sur une version d’évaluation API, utilisez le mécanisme de commentaires à la fin de la page web où l’API est documenté.

> [!NOTE]
> L’aperçu API peut être modifiés et n’est pas destinés à utiliser dans un environnement de production.

Pour plus d’informations sur les API de préversion, reportez-vous à l’article relatif à l’[ensemble de conditions requises de l’API Outlook de préversion](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md).
