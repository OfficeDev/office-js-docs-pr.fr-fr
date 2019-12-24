---
title: Ensembles de conditions requises de l’API JavaScript pour Excel
description: Informations sur la configuration requise pour le complément Office sur les builds Excel
ms.date: 12/24/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: ce3a51c37128627765b587d1baef35d051c52917
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851375"
---
# <a name="excel-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour Excel

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

## <a name="requirement-set-availability"></a>Disponibilité d’ensemble de conditions requises

Les compléments Excel peuvent être exécutés dans différentes versions d’Office, notamment Office 2016 ou version ultérieure sur Windows, et Office sur le web, iPad et Mac. Le tableau suivant répertorie les ensembles de conditions requises pour Excel, les applications hôtes Office qui prennent en charge chaque ensemble de conditions et la version ou le numéro de build de ces applications.

> [!NOTE]
> Pour utiliser des API dans l’un des ensembles de conditions requises numérotés ou `ExcelApiOnline`, vous devez référencer la bibliothèque de **production** sur le CDN : https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> Pour plus d’informations sur l’utilisation aperçu API, voir l’article[JavaScript d’Excel preview API](./excel-preview-apis.md).

|  Ensemble de conditions requises  |  Office pour Windows<br>(connecté à l’abonnement Office 365)  |  Office sur iPad<br>(connecté à l’abonnement Office 365)  |  Office sur Mac<br>(connecté à l’abonnement Office 365)  | Office sur le web |
|:-----|-----|:-----|:-----|:-----|:-----|
| [Aperçu](excel-preview-apis.md)  | Veuillez utiliser la dernière version d’Office pour tester la préversion API (vous devrez peut-être rejoindre la [programme Office Insider](https://products.office.com/office-insider)) |
| [ExcelApiOnline](excel-api-online-requirement-set.md) | N/A | N/A | N/A | Dernière version (voir la [page des ensembles de conditions requises](./excel-api-online-requirement-set.md)) |
| [ExcelApi 1.10](excel-api-1-10-requirement-set.md) | Version 1907 (Build 11929.20306) ou version ultérieure | 2.30 ou version ultérieure | 16.30 ou version ultérieure | Octobre 2019 |
| [ExcelApi 1.9](excel-api-1-9-requirement-set.md)  | Version 1903 (Build 11425.20204) ou version ultérieure | 2.24 ou version ultérieure | 16.24 ou version ultérieure | Mai 2019 |
| [ExcelApi 1.8](excel-api-1-8-requirement-set.md)  | Version 1808 (build 10730.20102) ou ultérieure | 2.17 ou version ultérieure | 16.17 ou version ultérieure | Septembre 2018 |
| [ExcelApi 1.7](excel-api-1-7-requirement-set.md)  | Version 1801 (build 9001.2171) ou ultérieure   | 2.9 ou version ultérieure  | 16.9 ou version ultérieure  | Avril 2018 |
| [ExcelApi 1.6](excel-api-1-6-requirement-set.md)  | Version 1704 (Build 8201.2001) ou version ultérieure   | 2.2 ou version ultérieure  | 15.36 ou version ultérieure | Avril 2017 |
| [ExcelApi 1.5](excel-api-1-5-requirement-set.md)  | Version 1703 (Build 8067.2070) ou version ultérieure   | 2.2 ou version ultérieure  | 15.36 ou version ultérieure | Mars 2017 |
| [ExcelApi 1.4](excel-api-1-4-requirement-set.md)  | Version 1701 (Build 7870.2024) ou version ultérieure   | 2.2 ou version ultérieure  | 15.36 ou version ultérieure | Janvier 2017 |
| [ExcelApi 1.3](excel-api-1-3-requirement-set.md)  | Version 1608 (Build 7369.2055) ou version ultérieure   | 1.27 ou version ultérieure | 15.27 ou version ultérieure | Septembre 2016 |
| [ExcelApi 1.2](excel-api-1-2-requirement-set.md)  | Version 1601 (Build 6741.2088) ou version ultérieure   | 1.21 ou version ultérieure | 15.22 ou version ultérieure | Janvier 2016 |
| [ExcelApi 1.1](excel-api-1-1-requirement-set.md)  | Version 1509 (Build 4266.1001) ou version ultérieure   | 1.19 ou version ultérieure | 15.20 ou version ultérieure | Janvier 2016 |

> [!NOTE]
> Le numéro de build d’Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que l’ensemble de conditions requises de l’ExcelApi 1.1.

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour plus d’informations sur les versions et les numéros de build d’Office, voir :

- [Numéros de version et de build des canaux de réception des mises à jour pour les clients Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Où trouver le numéro de version et de build pour une application cliente Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel)
- [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Spécification des exigences en matière d’hôtes Office et d’API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifeste XML des compléments Office](/office/dev/add-ins/develop/add-in-manifests)
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)
