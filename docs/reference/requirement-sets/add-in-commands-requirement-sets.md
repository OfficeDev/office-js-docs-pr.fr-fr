---
title: Ensembles de conditions requises concernant les commandes de complément
description: Vue d’ensemble des ensembles de conditions requises pour les commandes de complément Office
ms.date: 03/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: d1904d092988a445be3e481123eecbad39097764
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094413"
---
# <a name="add-in-commands-requirement-sets"></a>Ensembles de conditions requises concernant les commandes de complément

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. For more information, see [Add-in commands for Excel, Word, and PowerPoint](../../design/add-in-commands.md) and [Add-in commands for Outlook](../../outlook/add-in-commands-for-outlook.md).

The initial release of add-in commands doesn't have a corresponding requirement set (that is, there isn't an AddinCommands 1.0 requirement set). The following table lists the Office host applications that support the initial release version, and the build versions or number for those applications.  

| Version   |  Office 2013 sur Windows<br>(achat définitif) | Office 2016 sur Windows<br>(achat définitif) | Office 2019 sur Windows<br>(achat définitif) | Office pour Windows<br>(connecté à l’abonnement Microsoft 365)   |  Office sur iPad<br>(connecté à l’abonnement Microsoft 365)  |  Office sur Mac<br>(connecté à l’abonnement Microsoft 365)  | Office sur le web  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Commandes de complément (version initiale, aucune condition) | S/O | 16.0.4678.1000 *Pris en charge uniquement dans Outlook* | Version 1809 (build 10827.20150) ou version ultérieure |Version 1603 (build 6769.0000) ou ultérieure | S/O | 15.33 ou version ultérieure| Janvier 2016 |

L’ensemble de conditions de la version 1.1 des commandes de complément présente la possibilité d’[ouvrir automatiquement un volet de tâches avec des documents](../../develop/automatically-open-a-task-pane-with-a-document.md).

Le tableau suivant répertorie les ensembles de conditions requises des commandes de complément 1.1, les applications Office hôtes qui prennent en charge ces conditions et les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  |  Office 2013 sur Windows<br>(achat définitif) | Office 2016 sur Windows<br>(achat définitif) | Office 2019 sur Windows<br>(achat définitif) | Office pour Windows<br>(connecté à l’abonnement Microsoft 365)   |  Office sur iPad<br>(connecté à l’abonnement Microsoft 365)  |  Office sur Mac<br>(connecté à l’abonnement Microsoft 365)  | Office sur le web  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1.1  | S/O | 16.0.4678.1000 *Pris en charge uniquement dans Outlook*  | Version 1809 (build 10827.20150) ou version ultérieure | Version 1705 (build 8121.1000) ou ultérieure | S/O | 15.34 ou version ultérieure\*| Mai 2017 |

>\* La méthode [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) renverra `false` par erreur pour les versions 16.9 &ndash; 16.14 (incluse), mais l’ensemble de conditions requises *est* pris en charge sur ces versions.

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécification des exigences en matière d’hôtes Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
