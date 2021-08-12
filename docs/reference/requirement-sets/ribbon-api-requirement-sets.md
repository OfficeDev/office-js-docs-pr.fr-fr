---
title: Ensembles de conditions requises des API ruban
description: Spécifie les plateformes Office et les builds qui prisent en charge les API du ruban dynamique.
ms.date: 05/12/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 18b081da28991eb15875c6999ddfe06ccea13eade8dd8a4e1f2a4cc09ac5e922
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084922"
---
# <a name="ribbon-api-requirement-sets"></a>Ensembles de conditions requises des API ruban

Les ensembles de conditions requises sont des groupes nommés de membres de l’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’un complément nécessite. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

L’ensemble d’API du Ruban prend en charge le contrôle par programme du moment où les commandes de module personnalisées (c’est-à-dire, les boutons de ruban personnalisés et les éléments de menu) sont activées et désactivées.

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises de l’API du ruban, les applications clientes Office qui la prise en charge, ainsi que les numéros de build ou de version de l’application Office client.

|  Ensemble de conditions requises  | Office 2013 sur Windows<br>(achat définitif) | Office 2016 ou une Windows<br>(achat définitif)   | Office pour Windows\*<br>(connecté à un abonnement Microsoft 365) |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac\*<br>(connecté à un abonnement Microsoft 365)  | Office sur le web\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonAPI 1.1  | N/A | N/A | Voir la prise en charge<br>section ci-dessous | N/A | 16.38 | Novembre 2020 | N/A|
| RibbonApi 1.2  | N/A | N/A | 2102 (build 13801.20294) | N/A | bientôt disponible | Mai 2021 | N/A|

> **&#42;** L’API ruban est prise en charge uniquement sur Excel et nécessite un abonnement Microsoft 365 de connexion.

## <a name="support-for-version-11-on-office-on-windows-subscription"></a>Prise en charge de la version 1.1 Office sur Windows (abonnement)

La version 1.1 de l’ensemble de conditions requises RibbonApi est prise en charge dans le Canal consommateur version 2006 (build 13001.20498 ou version supérieure). Pour Office sur Windows la fonctionnalité est également prise en charge dans les builds du canal Semi-Annual et du canal Enterprise mensuel disponibles le 14 juillet 2020 ou une date ultérieure. Les builds minimales prise en charge pour chaque canal sont les suivantes :  

|Canal | Version | Build|
|:-----|:-----|:-----|
|Canal actuel | 2006 ou supérieure | 20266.20266 ou supérieur|
|Canal Entreprise mensuel | 2005 ou supérieure | 12827.20538 ou supérieur|
|Canal Entreprise mensuel | 2004 | 12730.20602 ou supérieur|
|Canal Entreprise semestriel | 2002 ou supérieure | 12527.20880 ou supérieur|

## <a name="more-information"></a>Informations supplémentaires

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

- [Numéros de version et de build des versions de canal de mise à jour Microsoft 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Où trouver la version et le numéro de build d’une application Microsoft 365 client](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="ribbon-api-11"></a>API de ruban 1.1

L’API du Ruban 1.1 est la première version de l’API. Pour plus d’informations sur l’API, voir [la rubrique Office.ribbon.](/javascript/api/office/office.ribbon)

## <a name="ribbon-api-12"></a>API de ruban 1.2

L’API 1.2 du Ruban ajoute la prise en charge des onglets contextuels. Si vous souhaitez en savoir, veuillez consulter la rubrique [Créer des onglets contextuels personnalisés dans des compléments Office](../../design/contextual-tabs.md).

> [!NOTE]
> L’ensemble de conditions requises **RibbonApi 1.2** n’est pas encore pris en charge dans le manifeste. Vous ne devez donc pas le spécifier dans la section du `<Requirements>` manifeste.

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécifier les exigences en matière d’applications Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
