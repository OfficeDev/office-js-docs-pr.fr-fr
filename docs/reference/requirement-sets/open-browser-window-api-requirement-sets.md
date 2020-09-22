---
title: Séries de conditions requises pour ouvrir une fenêtre de navigateur
description: Spécifie les plateformes et les générations Office qui prennent en charge l’API openBrowserWindow.
ms.date: 09/16/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 8bc26525bf64ed87d46d85cd1248f79696d67f2b
ms.sourcegitcommit: 4a03d8b3f676ee2d91114813cb81bce5da3c8d6b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/22/2020
ms.locfileid: "48175506"
---
# <a name="open-browser-window-api-requirement-sets"></a>Ouvrir les ensembles de conditions requises de l’API de fenêtre de navigateur

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

L’ensemble d’API OpenBrowserWindow permet à des compléments d’ouvrir un navigateur pour accomplir des tâches qui ne peuvent pas toujours être effectuées dans le contrôle de WebView en mode bac à sable dans le complément lui-même ; par exemple, en téléchargeant un fichier PDF lorsque le contrôle WebView est fourni par Microsoft Edge.

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises de l’API OpenBrowserWindow, les applications hôtes Office qui prennent en charge ces conditions et les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  | Office 2013 sur Windows ou version ultérieure<br>(achat définitif) | Office pour Windows<br>(connecté à l’abonnement Office 365) |  Office sur iPad<br>(connecté à l’abonnement Office 365)  |  Office sur Mac<br>(connecté à l’abonnement Office 365)  | Office sur le web  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| OpenBrowserWindowApi 1,1  | S/O | Version 1810 (Build 16.0.11001.20074) ou version ultérieure | 16.0.0.0 ou version ultérieure | 16.0.0.0 ou version ultérieure | N/A | N/A|

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

- [Numéros de version et de build des canaux de réception des mises à jour pour les clients Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Où trouver le numéro de version et de build pour une application cliente Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="openbrowserwindowapi-11"></a>OpenBrowserWindowApi 1,1

Le OpenBrowserWindowApi 1,1 est la première version de l’API. Pour plus d’informations sur l’API, voir la rubrique de référence [Office. Context. UI](/javascript/api/office/office.context#ui) .

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécification des exigences en matière d’hôtes Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
