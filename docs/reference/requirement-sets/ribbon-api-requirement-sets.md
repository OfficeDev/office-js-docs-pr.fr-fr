---
title: Ensembles de conditions requises des API ruban
description: Spécifie les plateformes et builds Office qui peuvent supporter les API du ruban dynamique.
ms.date: 11/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 91c909755779d122fba8d77dc246784f6a0dd1a3
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839984"
---
# <a name="ribbon-api-requirement-sets"></a>Ensembles de conditions requises des API ruban

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’ils nécessitent. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

L’ensemble d’API du ruban prend en charge le contrôle par programme du moment où les commandes de module personnalisées (c’est-à-dire, les boutons de ruban personnalisés et les éléments de menu) sont activées et désactivées.

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises de l’API du Ruban, les applications clientes Office qui le supportent, ainsi que les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  | Office 2013 sur Windows<br>(achat définitif) | Office 2016 ou une édition ultérieure sur Windows<br>(achat définitif)   | Office pour Windows\*<br>(connecté à un abonnement Microsoft 365) |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac\*<br>(connecté à un abonnement Microsoft 365)  | Office sur le web\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.1  | N/A | N/A | Voir la prise en charge<br>section ci-dessous | S/O | 16.38 | Novembre 2020 | S/O|

> **&#42;** L’API ruban est prise en charge uniquement sur Excel et nécessite un abonnement Microsoft 365.

## <a name="office-on-windows-subscription-support"></a>Prise en charge d’Office sur Windows (abonnement)

L’ensemble de conditions requises est pris en charge dans le canal grand public version 2006 (build, 13001.20498 ou version supérieure). Pour Office sur Windows, la fonctionnalité est également prise en charge dans les builds canal Semi-Annual et Canal entreprise mensuel disponibles le 14 juillet 2020 ou une date ultérieure. Les builds minimales prise en charge pour chaque canal sont les suivantes :  

|Canal | Version | Build|
|:-----|:-----|:-----|
|Canal actuel | 2006 ou supérieure | 20266.20266 ou supérieur|
|Canal Entreprise mensuel | 2005 ou supérieure | 12827.20538 ou supérieur|
|Canal Entreprise mensuel | 2004 | 12730.20602 ou supérieur|
|Canal d’entreprise semi-annuel | 2002 ou supérieure | 12527.20880 ou supérieur|

## <a name="more-information"></a>Informations supplémentaires

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

- [Numéros de version et de build des versions de canal de mise à jour pour les clients Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Où trouver la version et le numéro de build d’une application cliente Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

> [!NOTE]
> L’ensemble de conditions requises **RibbonApi 1.1** n’est pas encore pris en charge dans le manifeste, vous ne pouvez donc pas le spécifier dans la section du `<Requirements>` manifeste.


## <a name="office-common-api-requirement-sets"></a>Séries de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="ribbon-api-11"></a>API de ruban 1.1

L’API du Ruban 1.1 est la première version de l’API. Pour plus d’informations sur l’API, voir la rubrique de référence [Office.ribbon.](/javascript/api/office/office.ribbon)

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécifier les applications Office et les exigences de l’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)