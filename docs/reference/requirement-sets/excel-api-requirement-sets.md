---
title: Ensembles de conditions requises de l’API JavaScript pour Excel
description: Informations sur la configuration requise pour le complément Office sur les builds Excel.
ms.date: 05/05/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 6fb5587b7eb3120a1e4b7db7dc6327bdcadc6691
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937150"
---
# <a name="excel-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour Excel

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’ils nécessitent. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

## <a name="requirement-set-availability"></a>Disponibilité d’ensemble de conditions requises

Les compléments Excel peuvent être exécutés dans différentes versions d’Office, notamment Office 2016 ou version ultérieure pour Windows, Office pour iPad, Office pour Mac et Office Online. Le tableau suivant répertorie les ensembles de conditions requises pour Excel, les applications clientes Office qui prennent en charge chaque ensemble de conditions et les versions ou numéro de build de ces applications.

> [!NOTE]
> Pour utiliser des API dans l’un des ensembles de conditions requises numérotés ou `ExcelApiOnline`, vous devez référencer la bibliothèque de **production** sur le CDN : https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> Pour plus d’informations sur l’utilisation aperçu API, voir l’article[JavaScript d’Excel preview API](excel-preview-apis.md).

|  Ensemble de conditions requises  |  Office pour Windows<br>(connecté à un abonnement Microsoft 365)  |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac<br>(connecté à un abonnement Microsoft 365)  | Office sur le web |
|:-----|-----|:-----|:-----|:-----|:-----|
| [Aperçu](excel-preview-apis.md)  | Veuillez utiliser la dernière version d’Office pour tester les API d’aperçu (vous devrez peut-être adhérer au [programme Office Insider](https://insider.office.com)). |
| [ExcelApiOnline](excel-api-online-requirement-set.md) | N/A | N/A | N/A | Dernière version (voir la [page des ensembles de conditions requises](excel-api-online-requirement-set.md)) |
| [ExcelApi 1.12](excel-api-1-12-requirement-set.md) | Version 2008 (Build 13127.20408) ou version ultérieure | 16.40 ou version ultérieure | 16.40 ou version ultérieure | Septembre 2020 |
| [ExcelApi 1.11](excel-api-1-11-requirement-set.md) | Version 2002 (Build 12527.20470) ou version ultérieure | 16.35 ou version ultérieure | 16.33 ou version ultérieure | Mai 2020 |
| [ExcelApi 1.10](excel-api-1-10-requirement-set.md) | Version 1907 (Build 11929.20306) ou version ultérieure | 16.0 ou version ultérieure | 16.30 ou version ultérieure | Octobre 2019 |
| [ExcelApi 1.9](excel-api-1-9-requirement-set.md)  | Version 1903 (Build 11425.20204) ou version ultérieure | 16.0 ou version ultérieure | 16.24 ou version ultérieure | Mai 2019 |
| [ExcelApi 1.8](excel-api-1-8-requirement-set.md)  | Version 1808 (build 10730.20102) ou ultérieure | 16.0 ou version ultérieure | 16.17 ou version ultérieure | Septembre 2018 |
| [ExcelApi 1.7](excel-api-1-7-requirement-set.md)  | Version 1801 (build 9001.2171) ou ultérieure   | 16.0 ou version ultérieure  | 16.9 ou version ultérieure  | Avril 2018 |
| [ExcelApi 1.6](excel-api-1-6-requirement-set.md)  | Version 1704 (Build 8201.2001) ou version ultérieure   | 15.0 ou version ultérieure  | 15.36 ou version ultérieure | Avril 2017 |
| [ExcelApi 1.5](excel-api-1-5-requirement-set.md)  | Version 1703 (Build 8067.2070) ou version ultérieure   | 15.0 ou version ultérieure  | 15.36 ou version ultérieure | Mars 2017 |
| [ExcelApi 1.4](excel-api-1-4-requirement-set.md)  | Version 1701 (Build 7870.2024) ou version ultérieure   | 15.0 ou version ultérieure  | 15.36 ou version ultérieure | Janvier 2017 |
| [ExcelApi 1.3](excel-api-1-3-requirement-set.md)  | Version 1608 (Build 7369.2055) ou version ultérieure   | 15.0 ou version ultérieure | 15.27 ou version ultérieure | Septembre 2016 |
| [ExcelApi 1.2](excel-api-1-2-requirement-set.md)  | Version 1601 (Build 6741.2088) ou version ultérieure   | 15.0 ou version ultérieure | 15.22 ou version ultérieure | Janvier 2016 |
| [ExcelApi 1.1](excel-api-1-1-requirement-set.md)  | Version 1509 (Build 4266.1001) ou version ultérieure   | 15.0 ou version ultérieure | 15.20 ou version ultérieure | Janvier 2016 |

> [!NOTE]
> Les versions sans abonnement d'Office prennent en charge l'ensemble des conditions requises suivantes :
>
> - Office 2019 prend en charge ExcelApi 1.8 et versions antérieures.
> - Office 2016 prend uniquement en charge l'ensemble des conditions requises de ExcelApi 1.1.

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour plus d’informations sur les versions et les numéros de build d’Office, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="how-to-use-excel-requirement-sets-at-runtime-and-in-the-manifest"></a>Utiliser les conditions requises Excel au moment de l’exécution et dans le manifeste

> [!NOTE]
> Cette section suppose que vous êtes familiarisé avec les rubriques [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md) et [Spécifier les applications Office et les exigences de l’API](../../develop/specify-office-hosts-and-api-requirements.md).

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Le complément Office peut effectuer une vérification à l’exécution ou utiliser des ensembles de conditions requises spécifiés dans le manifeste pour déterminer si une application Office prend en charge les API requises par le complément.

### <a name="checking-for-requirement-set-support-at-runtime"></a>Vérification de la prise en charge de l’ensemble de conditions requises à l’exécution

L’exemple de code suivant montre comment déterminer si l’application Office dans laquelle le complément est en cours d’exécution prend en charge l’ensemble spécifié de conditions requises pour l’API.

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>Définition de la prise en charge de l’ensemble de conditions requises dans le manifeste

Vous pouvez utiliser l’élément [Requirements](../manifest/requirements.md) dans le manifeste du complément pour spécifier les ensembles de conditions requises minimales et/ou les méthodes d’API nécessaires à l’activation de votre complément. Si l’application ou la plateforme Office ne prend pas en charge les ensembles de conditions requises ou les méthodes d’API spécifiés dans l’élément `Requirements` du manifeste, le complément ne s’exécute pas dans cette application ou cette plateforme et ne s’affiche pas dans la liste des compléments affichés dans **Mes compléments**. Si votre complément nécessite un ensemble de conditions requises spécifique pour toutes les fonctionnalités, mais qu’il peut fournir de la valeur même aux utilisateurs sur des plateformes qui ne prennent pas en charge l’ensemble de conditions requises, nous vous recommandons de vérifier la prise en charge des exigences au moment de l’exécution, comme décrit ci-dessus, au lieu de définir la prise en charge de l’ensemble de conditions requises dans le manifeste.

L’exemple de code suivant montre l’élément `Requirements` dans un manifeste indiquant que le complément doit être chargé dans toutes les applications clientes Office prenant en charge l’ensemble de conditions requises ExcelApi version 1.3 ou ultérieure.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
