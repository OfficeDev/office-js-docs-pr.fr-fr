---
title: Ensembles de conditions requises de l’API JavaScript pour Excel
description: Informations sur la configuration requise pour le complément Office sur les builds Excel
ms.date: 03/11/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: b6e1570d7487e552197201d12f9a783f18a30fe3
ms.sourcegitcommit: 05b73cdec5f4db7f0b8d48a5a552ee296a0332ca
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42600703"
---
# <a name="excel-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour Excel

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

## <a name="requirement-set-availability"></a>Disponibilité d’ensemble de conditions requises

Les compléments Excel peuvent être exécutés dans différentes versions d’Office, notamment Office 2016 ou version ultérieure sur Windows, et Office sur le web, iPad et Mac. Le tableau suivant répertorie les ensembles de conditions requises pour Excel, les applications hôtes Office qui prennent en charge chaque ensemble de conditions et la version ou le numéro de build de ces applications.

> [!NOTE]
> Pour utiliser des API dans l’un des ensembles de conditions requises numérotés ou `ExcelApiOnline`, vous devez référencer la bibliothèque de **production** sur le CDN : https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> Pour plus d’informations sur l’utilisation aperçu API, voir l’article[JavaScript d’Excel preview API](excel-preview-apis.md).

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
> Les versions perpétuelles d'Office prennent en charge l'ensemble des conditions requises suivantes :
>
> - Office 2019 prend en charge ExcelApi 1.8 et versions antérieures.
> - Office 2016 prend uniquement en charge l'ensemble des conditions requises de ExcelApi 1.1.

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour plus d’informations sur les versions et les numéros de build d’Office, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="see-also"></a>Articles associés

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel)
- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécification des exigences en matière d’hôtes Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)
