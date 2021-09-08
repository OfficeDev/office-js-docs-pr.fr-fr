---
title: Ensembles de conditions requises concernant les commandes de complément
description: Vue d’Office ensembles de conditions requises des commandes de l’autre.
ms.date: 11/01/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: f5f7c07f9bdb6bee923337dcc2ae547ca1f76df3
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937277"
---
# <a name="add-in-commands-requirement-sets"></a>Ensembles de conditions requises concernant les commandes de complément

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’ils nécessitent. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Les commandes de complément sont des éléments d’interface utilisateur qui étendent l’interface utilisateur d’Office et lancent des actions dans votre complément. Vous pouvez les utiliser pour ajouter un bouton sur le ruban ou un élément dans le menu contextuel. Pour plus d’informations, reportez-vous à la rubrique sur les [commandes de complément pour Excel, Word et PowerPoint](../../design/add-in-commands.md) et celle sur les [commandes de complément pour Outlook](../../outlook/add-in-commands-for-outlook.md).

La version initiale des commandes de add-in n’a pas d’ensemble de conditions requises correspondant (autrement dit, il n’existe pas d’ensemble de conditions requises AddinCommands 1.0). Le tableau suivant répertorie les applications clientes Office qui la prise en charge de la version initiale, ainsi que les versions ou le numéro de build de ces applications.  

| Version   |  Office 2013 sur Windows<br>(achat définitif) | Office 2016 sur Windows<br>(achat définitif) | Office 2019 sur Windows<br>(achat définitif) | Office pour Windows<br>(connecté à un abonnement Microsoft 365)   |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac<br>(connecté à un abonnement Microsoft 365)  | Office sur le web  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Commandes de complément (version initiale, aucune condition) | S/O | 16.0.4678.1000 *Pris en charge uniquement dans Outlook* | Version 1809 (build 10827.20150) ou version ultérieure |Version 1603 (build 6769.0000) ou ultérieure | S/O | 15.33 ou version ultérieure| Janvier 2016 |

L’ensemble de conditions requises des commandes de add-in **1.1** introduit la possibilité d’ouverture automatique d’un volet De tâches [avec des documents.](../../develop/automatically-open-a-task-pane-with-a-document.md)

L’ensemble de conditions requises des commandes de l’ajout **1.3** introduit un marques de manifeste qui permet à un module de personnaliser l’emplacement d’un onglet personnalisé sur le ruban Office et d’insérer des contrôles de ruban Office intégrés dans des groupes de contrôles personnalisés.

Le tableau suivant répertorie les ensembles de conditions requises pour les commandes de Office, les applications clientes Office qui la prise en charge, ainsi que les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  |  Office 2013 sur Windows<br>(achat définitif) | Office 2016 sur Windows<br>(achat définitif) | Office 2019 sur Windows<br>(achat définitif) | Office pour Windows<br>(connecté à un abonnement Microsoft 365)   |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac<br>(connecté à un abonnement Microsoft 365)  | Office sur le web  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1.3  | N/A | N/A  | N/A | bientôt disponible | N/A | bientôt disponible | Novembre 2020 |
| AddinCommands 1.1  | S/O | 16.0.4678.1000 *Pris en charge uniquement dans Outlook*  | Version 1809 (build 10827.20150) ou version ultérieure | Version 1705 (build 8121.1000) ou ultérieure | S/O | 15.34 ou version ultérieure\*| Mai 2017 |

>\* La méthode [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#isSetSupported_name__minVersion_) renverra `false` par erreur pour les versions 16.9 &ndash; 16.14 (incluse), mais l’ensemble de conditions requises *est* pris en charge sur ces versions.

> [!IMPORTANT]
> AddinCommands 1.3 est en prévisualisation et n’est disponible que *dans PowerPoint sur le web*. Nous vous recommandons d’essayer le markup uniquement dans les environnements de test et de développement. N’utilisez pas de marques d’aperçu dans un environnement de production ou dans des documents critiques pour l’entreprise.

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécifier les exigences en matière d’applications Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
