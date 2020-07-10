---
title: Ensembles de conditions requises de l’API de dialogue
description: En savoir plus sur les ensembles de conditions requises de l’API Dialog.
ms.date: 06/10/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: d50c30fd769777c8dd3c168a9289dfb60012bbbd
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094427"
---
# <a name="dialog-api-requirement-sets"></a>Ensembles de conditions requises de l’API de dialogue

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Office Add-ins run across multiple versions of Office. The following table lists the Dialog API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.

|  Ensemble de conditions requises  | Office 2013 sur Windows\*<br>(achat définitif) | Office 2016 ou version ultérieure sur Windows\*<br>(achat définitif)   | Office pour Windows<br>(connecté à l’abonnement Microsoft 365) |  Office sur iPad<br>(connecté à l’abonnement Microsoft 365)  |  Office sur Mac<br>(connecté à l’abonnement Microsoft 365)  | Office sur le web  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | Build 15.0.4855.1000 ou version ultérieure | Build 16.0.4390.1000 ou version ultérieure | Version 1602 (Build 6741.0000) ou version ultérieure | 1.22 ou version ultérieure | 15.20 ou version ultérieure| Janvier 2017 | Version 1608 (Build 7601.6800) ou version ultérieure|

>\*Les utilisateurs du bureau unique peuvent ne pas avoir accepté tous les correctifs et mises à jour. Si c’est le cas, la DLL qu’Office utilise pour signaler sa version dans l’interface utilisateur peut être supérieure aux versions indiquées ici même si les dll mises à jour nécessaires pour prendre en charge DialogApi n’ont pas été installées sur l’ordinateur de l’utilisateur. Pour vous assurer que le correctif nécessaire est installé, l’utilisateur doit accéder à la liste Office Update List ([office 2013 List](/officeupdates/msp-files-office-2013) ou [Office 2016 List](/officeupdates/msp-files-office-2016)), rechercher **osfclient-x-none**et installer le correctif répertorié.

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="dialog-api-11"></a>API de boîte de dialogue 1.1

L’API de boîte de dialogue 1.1 est la première version de l’API. Pour plus d’informations sur l’API, consultez la rubrique référence de l' [API Dialog](/javascript/api/office/office.ui) .

## <a name="see-also"></a>Voir aussi

- [Utiliser l’API de boîte de dialogue Office dans les compléments Office](../../develop/dialog-api-in-office-add-ins.md)
- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécification des exigences en matière d’hôtes Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
