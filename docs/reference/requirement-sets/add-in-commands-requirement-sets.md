---
title: Ensembles de conditions requises concernant les commandes de complément
description: ''
ms.date: 06/20/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c6c71e01dff2c8bc595d662e5897a4c98692a216
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163954"
---
# <a name="add-in-commands-requirement-sets"></a>Ensembles de conditions requises concernant les commandes de complément

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Les commandes de complément sont des éléments d’interface utilisateur qui étendent l’interface utilisateur d’Office et lancent des actions dans votre complément. Vous pouvez les utiliser pour ajouter un bouton sur le ruban ou un élément dans le menu contextuel. Pour plus d’informations, reportez-vous à la rubrique sur les [commandes de complément pour Excel, Word et PowerPoint](/office/dev/add-ins/design/add-in-commands) et celle sur les [commandes de complément pour Outlook](../../outlook/add-in-commands-for-outlook.md).

Il n’existe pas d’ensemble de conditions particulier pour la version initiale des commandes de complément (autrement dit, il n’existe pas d’ensemble de conditions AddInCommands 1.0). Le tableau suivant présente les applications hôtes Office qui prennent en charge la version initiale, ainsi que leur build ou leur numéro de version.  

| Version   |  Office 2013 sur Windows<br>(achat définitif) | Office 2016 sur Windows<br>(achat définitif) | Office 2019 sur Windows<br>(achat définitif) | Office pour Windows<br>(connecté à l’abonnement Office 365)   |  Office sur iPad<br>(connecté à l’abonnement Office 365)  |  Office sur Mac<br>(connecté à l’abonnement Office 365)  | Office sur le web  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Commandes de complément (version initiale, aucune condition) | S/O | 16.0.4678.1000 *Pris en charge uniquement dans Outlook* | Version 1809 (build 10827.20150) ou version ultérieure |Version 1603 (build 6769.0000) ou ultérieure | S/O | 15.33 ou version ultérieure| Janvier 2016 |

L’ensemble de conditions de la version 1.1 des commandes de complément présente la possibilité d’[ouvrir automatiquement un volet de tâches avec des documents](/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document).

Le tableau suivant répertorie les ensembles de conditions requises des commandes de complément 1.1, les applications Office hôtes qui prennent en charge ces conditions et les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  |  Office 2013 sur Windows<br>(achat définitif) | Office 2016 sur Windows<br>(achat définitif) | Office 2019 sur Windows<br>(achat définitif) | Office pour Windows<br>(connecté à l’abonnement Office 365)   |  Office sur iPad<br>(connecté à l’abonnement Office 365)  |  Office sur Mac<br>(connecté à l’abonnement Office 365)  | Office sur le web  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1.1  | S/O | 16.0.4678.1000 *Pris en charge uniquement dans Outlook*  | Version 1809 (build 10827.20150) ou version ultérieure | Version 1705 (build 8121.1000) ou ultérieure | S/O | 15.34 ou version ultérieure\*| Mai 2017 |

>\* La méthode [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) renverra `false` par erreur pour les versions 16.9 &ndash; 16.14 (incluse), mais l’ensemble de conditions requises *est* pris en charge sur ces versions.

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

- [Numéros de version et de build des canaux de réception des mises à jour pour les clients Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Où trouver le numéro de version et de build pour une application cliente Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Spécification des exigences en matière d’hôtes Office et d’API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifeste XML des compléments Office](/office/dev/add-ins/develop/add-in-manifests)
