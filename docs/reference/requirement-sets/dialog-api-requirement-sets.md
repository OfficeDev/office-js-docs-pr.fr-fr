---
title: Ensembles de conditions requises de l’API de dialogue
description: ''
ms.date: 07/05/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: a524edf6734618a56e050d2c25eedbd23ca13973
ms.sourcegitcommit: 9c5a836d4464e49846c9795bf44cfe23e9fc8fbe
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2019
ms.locfileid: "35617020"
---
# <a name="dialog-api-requirement-sets"></a>Ensembles de conditions requises de l’API de dialogue

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises de l’API de boîte de dialogue, les applications Office hôte qui prennent en charge ces conditions et les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  | Office 2013 sur Windows\*<br>(achat définitif) | Office 2016 ou version ultérieure sur Windows\*<br>(achat définitif)   | Office sur Windows<br>(connecté à l’abonnement Office 365) |  Office sur iPad<br>(connecté à l’abonnement Office 365)  |  Office sur Mac<br>(connecté à l’abonnement Office 365)  | Office sur le Web  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | Build 15.0.4855.1000 ou version ultérieure | Build 16.0.4390.1000 ou version ultérieure | Version 1602 (Build 6741.0000) ou version ultérieure | 1.22 ou version ultérieure | 15.20 ou version ultérieure| Janvier 2017 | Version 1608 (Build 7601.6800) ou version ultérieure|

>\*Les utilisateurs du bureau unique peuvent ne pas avoir accepté tous les correctifs et mises à jour. Si c’est le cas, la DLL qu’Office utilise pour signaler sa version dans l’interface utilisateur peut être supérieure aux versions indiquées ici même si les dll mises à jour nécessaires pour prendre en charge DialogApi n’ont pas été installées sur l’ordinateur de l’utilisateur. Pour vous assurer que le correctif nécessaire est installé, l’utilisateur doit accéder à la liste Office Update List ([office 2013 List](/officeupdates/msp-files-office-2013) ou [Office 2016 List](/officeupdates/msp-files-office-2016)), rechercher **osfclient-x-none**et installer le correctif répertorié. 

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

- [Numéros de version et de build des canaux de réception des mises à jour pour les clients Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Où trouver le numéro de version et de build pour une application cliente Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="dialog-api-11"></a>API de boîte de dialogue 1.1

L’API de boîte de dialogue 1.1 est la première version de l’API. Pour plus d’informations sur l’API, consultez les rubriques de référence sur l’[API de boîte de dialogue](/javascript/api/office/office.ui).

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Spécification des exigences en matière d’hôtes Office et d’API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifeste XML des compléments Office](/office/dev/add-ins/develop/add-in-manifests)
