---
title: Ensembles de conditions requises de l’API JavaScript pour PowerPoint
description: ''
ms.date: 07/26/2019
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 4f64654a4130cc0d4bf96d9c59e364e77c808748
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35941145"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour PowerPoint

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Le tableau suivant répertorie les ensembles de conditions requises pour PowerPoint, les applications hôtes Office qui prennent en charge ces ensembles de conditions requises, ainsi que les versions de build ou la date de disponibilité.

|  Ensemble de conditions requises  |  Office pour Windows<br>(connecté à l’abonnement Office 365)  |  Office sur iPad<br>(connecté à l’abonnement Office 365)  |  Office sur Mac<br>(connecté à l’abonnement Office 365)  | Office sur le Web |
|:-----|-----|:-----|:-----|:-----|:-----|
| PowerPointApi 1,1 | Version 1810 (Build 11001,20074) ou version ultérieure | 2.17 ou version ultérieure | 16,19 ou version ultérieure | Octobre 2018 |

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour plus d’informations sur les versions d’Office et les numéros de build, voir:

- [Numéros de version et de build des canaux de réception des mises à jour pour les clients Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Où trouver le numéro de version et de build pour une application cliente Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## <a name="powerpoint-javascript-api-11"></a>API JavaScript pour PowerPoint 1,1

L’API JavaScript pour PowerPoint 1,1 contient une seule API pour créer une nouvelle présentation. Pour plus d’informations sur l’API, consultez la rubrique [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).

## <a name="runtime-requirement-support-check"></a>Vérification de la prise en charge d’un ensemble de conditions requises à l’exécution

Lors de l’exécution, les compléments peuvent vérifier si un hôte particulier prend en charge un ensemble de conditions requises de l’API en procédant comme suit.

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>Vérification de la prise en charge des conditions requises basée sur le manifeste

Utilisez l' `Requirements` élément dans le manifeste du complément pour spécifier les ensembles de conditions requises critiques ou les membres de l’API que votre complément doit utiliser. Si l’hôte ou la plateforme Office ne prend pas en charge les ensembles de conditions requises `Requirements` ou les membres d’API spécifiés dans l’élément, le complément ne s’exécutera pas sur cet hôte ou cette plateforme, et ne s’affichera pas dans mes compléments.

Cet exemple de code illustre un complément qui se charge dans toutes les applications hôtes Office qui prennent en charge l’ensemble de conditions requises OneNoteApi, version 1.1.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

La plupart des fonctionnalités de complément PowerPoint proviennent de l’ensemble d’API commun. Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour PowerPoint](/javascript/api/powerpoint)
- [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Spécification des exigences en matière d’hôtes Office et d’API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifeste XML des compléments Office](/office/dev/add-ins/develop/add-in-manifests)
