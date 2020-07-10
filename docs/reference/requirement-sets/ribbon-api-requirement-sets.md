---
title: Ensembles de conditions requises des API ruban
description: Spécifie les plateformes et les générations Office qui prennent en charge les API Dynamic Ribbon.
ms.date: 07/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 6a0e6af3a74b0b0402710fd66bac6c915aa4c18a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094280"
---
# <a name="ribbon-api-requirement-sets"></a>Ensembles de conditions requises des API ruban

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

L’ensemble d’API du ruban prend en charge le contrôle par programme de lorsque des commandes de complément personnalisées (c’est-à-dire des boutons personnalisés du ruban et des éléments de menu) sont activées et désactivées.

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises de l’API du ruban, les applications hôtes Office qui prennent en charge l’ensemble de conditions requises, ainsi que les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  | Office 2013 sur Windows<br>(achat définitif) | Office 2016 ou version ultérieure sur Windows<br>(achat définitif)   | Office pour Windows\*<br>(connecté à l’abonnement Microsoft 365) |  Office sur iPad<br>(connecté à l’abonnement Microsoft 365)  |  Office sur Mac\*<br>(connecté à l’abonnement Microsoft 365)  | Office sur le web\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1,1  | N/A | N/A | Version 2002 (Build 12527,20264) ou version ultérieure | 16,38 ou version ultérieure | S/O | Février 2020 | S/O|

> **&#42;** Pendant la phase d’aperçu, l’API du ruban est prise en charge uniquement sur Excel et nécessite un abonnement Microsoft 365. Vous devez utiliser la version et le build mensuels les plus récents du canal du programme Insider. Vous devez participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://products.office.com/office-insider?tab=tab-1). Notez que lorsqu’une build est basée sur le canal semi-annuel de production, la prise en charge des fonctionnalités d’aperçu, y compris l’API du ruban, est désactivée pour cette version.

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

- [Numéros de version et de build des canaux de mise à jour pour les clients Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Où trouver le numéro de version et de build pour une application cliente Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="ribbon-api-11"></a>API de ruban 1,1

L’API de ruban 1,1 est la première version de l’API. Pour plus d’informations sur l’API, reportez-vous à la rubrique Référence du [ruban Office](/javascript/api/office/office.ribbon) .

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Spécification des exigences en matière d’hôtes Office et d’API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifeste XML des compléments Office](/office/dev/add-ins/develop/add-in-manifests)
