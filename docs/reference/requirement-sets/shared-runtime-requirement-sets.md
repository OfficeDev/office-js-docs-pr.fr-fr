---
title: Ensembles de conditions requises pour l’exécution partagée
description: Spécifie les plateformes et les hôtes Office qui prennent en charge les API SharedRuntime.
ms.date: 02/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: dbb9d908154da074eaff6901c778adea168504a9
ms.sourcegitcommit: 7464eac3b54a6a6b65e27549a3ad603af6ee1011
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42315879"
---
# <a name="shared-runtime-requirement-sets"></a>Ensembles de conditions requises pour l’exécution partagée

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Les parties d’un complément Office qui exécutent du code JavaScript, telles que des volets de tâches, des fichiers de fonctions lancés à partir de commandes de complément et des fonctions personnalisées Excel, peuvent partager un seul Runtime JavaScript. Cela permet à toutes les parties de partager un ensemble de variables globales, de partager un ensemble de bibliothèques chargées et de communiquer les uns avec les autres sans avoir à transmettre de messages via un stockage persistant.

Le tableau suivant répertorie l’ensemble de conditions requises SharedRuntime 1,1, les applications hôtes Office qui prennent en charge cet ensemble de conditions requises, ainsi que les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  |  Office 2013 (ou version ultérieure) sur Windows<br>(achat définitif) | Office pour Windows<br>(connecté à l’abonnement Office 365)   |  Office sur iPad<br>(connecté à l’abonnement Office 365)  |  Office sur Mac<br>(connecté à l’abonnement Office 365)  | Office sur le web  | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| SharedRuntime 1,1  | S/O | Version 2002 (Build 12527,20092) ou version ultérieure | S/O | 16,35 ou version ultérieure | Février 2020 | S/O |

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
