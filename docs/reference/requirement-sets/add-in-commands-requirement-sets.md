---
title: Ensembles de conditions requises concernant les commandes de complément
description: Vue d’ensemble des ensembles de conditions requises pour les commandes de complément Office.
ms.date: 11/01/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 08fcb5df0e614e9b9f3ec9479fc958cc79adf320
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087958"
---
# <a name="add-in-commands-requirement-sets"></a>Ensembles de conditions requises concernant les commandes de complément

Les ensembles de conditions requises sont des groupes nommés de membres de l’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’un complément nécessite. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Les commandes de complément sont des éléments d’interface utilisateur qui étendent l’interface utilisateur d’Office et lancent des actions dans votre complément. Vous pouvez les utiliser pour ajouter un bouton sur le ruban ou un élément dans le menu contextuel. Pour plus d’informations, reportez-vous à la rubrique sur les [commandes de complément pour Excel, Word et PowerPoint](../../design/add-in-commands.md) et celle sur les [commandes de complément pour Outlook](../../outlook/add-in-commands-for-outlook.md).

La version initiale des commandes de complément n’a pas d’ensemble de conditions requises correspondantes (autrement dit, il n’existe pas d’ensemble de conditions requises AddinCommands 1,0). Le tableau suivant répertorie les applications clientes Office qui prennent en charge la version initiale, ainsi que les versions ou le numéro de build de ces applications.  

| Version   |  Office 2013 sur Windows<br>(achat définitif) | Office 2016 sur Windows<br>(achat définitif) | Office 2019 sur Windows<br>(achat définitif) | Office pour Windows<br>(connecté à un abonnement Microsoft 365)   |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac<br>(connecté à un abonnement Microsoft 365)  | Office sur le web  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Commandes de complément (version initiale, aucune condition) | S/O | 16.0.4678.1000 *Pris en charge uniquement dans Outlook* | Version 1809 (build 10827.20150) ou version ultérieure |Version 1603 (build 6769.0000) ou ultérieure | S/O | 15.33 ou version ultérieure| Janvier 2016 |

L’ensemble de conditions requises pour les commandes de complément **1,1** offre la possibilité d' [ouvrir automatique un volet Office avec des documents](../../develop/automatically-open-a-task-pane-with-a-document.md).

L’ensemble de conditions requises pour les commandes de complément **1,3** introduit un balisage de manifeste qui permet à un complément de personnaliser le positionnement d’un onglet personnalisé sur le ruban Office et d’insérer des contrôles de ruban Office intégrés dans des groupes de contrôles personnalisés.

Le tableau suivant répertorie les ensembles de conditions requises pour les commandes de complément, les applications clientes Office qui prennent en charge cet ensemble de conditions requises, ainsi que les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  |  Office 2013 sur Windows<br>(achat définitif) | Office 2016 sur Windows<br>(achat définitif) | Office 2019 sur Windows<br>(achat définitif) | Office pour Windows<br>(connecté à un abonnement Microsoft 365)   |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac<br>(connecté à un abonnement Microsoft 365)  | Office sur le web  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1,3  | N/A | N/A  | S/O | bientôt disponible | S/O | bientôt disponible | Novembre 2020 |
| AddinCommands 1.1  | S/O | 16.0.4678.1000 *Pris en charge uniquement dans Outlook*  | Version 1809 (build 10827.20150) ou version ultérieure | Version 1705 (build 8121.1000) ou ultérieure | S/O | 15.34 ou version ultérieure\*| Mai 2017 |

>\* La méthode [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) renverra `false` par erreur pour les versions 16.9 &ndash; 16.14 (incluse), mais l’ensemble de conditions requises *est* pris en charge sur ces versions.

> [!IMPORTANT]
> AddinCommands 1,3 est en préversion et n’est *disponible que dans PowerPoint sur le Web*. Nous vous recommandons d’essayer le balisage uniquement dans les environnements de test et de développement. N’utilisez pas les marques de révision dans un environnement de production ou dans des documents professionnels.

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécifier les applications Office et les exigences de l’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
