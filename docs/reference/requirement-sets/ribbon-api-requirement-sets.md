---
title: Ensembles de conditions requises des API ruban
description: Spécifie les plateformes et les générations Office qui prennent en charge les API Dynamic Ribbon.
ms.date: 11/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 878670367b253fa7700434681244b43b9cfa36a7
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996514"
---
# <a name="ribbon-api-requirement-sets"></a>Ensembles de conditions requises des API ruban

Les ensembles de conditions requises sont des groupes nommés de membres de l’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’un complément nécessite. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

L’ensemble d’API du ruban prend en charge le contrôle par programme de lorsque des commandes de complément personnalisées (c’est-à-dire des boutons personnalisés du ruban et des éléments de menu) sont activées et désactivées.

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises de l’API du ruban, les applications clientes Office qui prennent en charge l’ensemble de conditions requises, ainsi que les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  | Office 2013 sur Windows<br>(achat définitif) | Office 2016 ou version ultérieure sur Windows<br>(achat définitif)   | Office pour Windows\*<br>(connecté à un abonnement Microsoft 365) |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac\*<br>(connecté à un abonnement Microsoft 365)  | Office sur le web\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1,1  | N/A | N/A | Voir prise en charge<br>section ci-dessous | N/A | 16,38 | Novembre 2020 | N/A|

> **&#42;** L’API du ruban est prise en charge uniquement sur Excel et nécessite un abonnement Microsoft 365.

## <a name="office-on-windows-subscription-support"></a>Prise en charge d’Office sur Windows (abonnement)

L’ensemble de conditions requises est pris en charge dans la version 2006 du canal grand public (Build, 13001,20498 ou version ultérieure). Pour Office sur Windows, la fonctionnalité est également prise en charge dans le canal Semi-Annual et les versions mensuelles de canaux d’entreprise disponibles pour le 14 juillet 2020 ou version ultérieure. Les versions minimales prises en charge pour chaque canal sont les suivantes :  

|Canal | Version | Build|
|:-----|:-----|:-----|
|Canal actuel | 2006 ou version ultérieure | 20266,20266 ou version ultérieure|
|Canal Entreprise mensuel | 2005 ou version ultérieure | 12827,20538 ou version ultérieure|
|Canal Entreprise mensuel | 2004 | 12730,20602 ou version ultérieure|
|Canal d’entreprise semi-annuel | 2002 ou version ultérieure | 12527,20880 ou version ultérieure|

## <a name="more-information"></a>Plus d’informations

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

- [Numéros de version et de build des canaux de réception des mises à jour pour les clients Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Où trouver le numéro de version et de build pour une application cliente Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

> [!NOTE]
> L’ensemble de conditions requises **RibbonApi 1,1** n’étant pas encore pris en charge dans le manifeste, vous ne pouvez pas le spécifier dans la section du manifeste `<Requirements>` .


## <a name="office-common-api-requirement-sets"></a>Séries de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="ribbon-api-11"></a>API de ruban 1,1

L’API de ruban 1,1 est la première version de l’API. Pour plus d’informations sur l’API, reportez-vous à la rubrique Référence du [ruban Office ](/javascript/api/office/office.ribbon) .

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Spécifier les applications Office et les exigences de l’API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifeste XML des compléments Office](/office/dev/add-ins/develop/add-in-manifests)
