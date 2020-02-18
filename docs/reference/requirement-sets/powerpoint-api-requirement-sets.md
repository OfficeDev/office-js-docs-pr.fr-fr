---
title: Séries de conditions requises de l’API JavaScript pour PowerPoint
description: ''
ms.date: 07/26/2019
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: 5bba2354cabba3c3ccd4ddf38d3e03c25a32b8a9
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950956"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a>Séries de conditions requises de l’API JavaScript pour PowerPoint

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Le tableau suivant répertorie les séries de conditions requises pour PowerPoint, les applications hôtes Office qui prennent en charge ces conditions et les numéros de version ou la date de disponibilité.

|  Ensemble de conditions requises  |  Office pour Windows<br>(connecté à l’abonnement Office 365)  |  Office sur iPad<br>(connecté à l’abonnement Office 365)  |  Office sur Mac<br>(connecté à l’abonnement Office 365)  | Office sur le web |
|:-----|-----|:-----|:-----|:-----|:-----|
| PowerPointApi 1.1 | Version 1810 (Build 11001.20074) ou version ultérieure | 2.17 ou version ultérieure | 16.19 ou version ultérieure | Octobre 2018 |

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour plus d’informations sur les versions et les numéros de build d’Office, voir :

- [Numéros de version et de build des canaux de réception des mises à jour pour les clients Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Où trouver le numéro de version et de build pour une application cliente Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## <a name="powerpoint-javascript-api-11"></a>API JavaScript pour PowerPoint 1.1

L’API JavaScript PowerPoint 1.1 inclut une seule API pour créer une nouvelle présentation. Pour plus de détails sur l’API, voir [API JavaScript pour PowerPoint](../../powerpoint/powerpoint-add-ins.md).

## <a name="runtime-requirement-support-check"></a>Vérification de la prise en charge d’une exigence d'exécution

Lors de l’exécution, les compléments peuvent vérifier si un hôte particulier prend en charge une série de conditions requises d’API en procédant comme suit.

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>Vérification de la prise en charge des conditions requises basées sur le manifeste

Utilisez l’élément `Requirements` dans le manifeste du complément pour spécifier des ensembles de conditions requises essentiels ou des membres d’API que votre complément doit utiliser. Si la plateforme ou l’hôte Office ne prend pas en charge les ensembles de conditions requises ou les membres d’API spécifiés dans l’élément `Requirements`, le complément ne s’exécute pas dans cet hôte ou cette plateforme et ne s’affiche pas dans Mes compléments.

Cet exemple de code illustre un complément qui se charge dans toutes les applications hôtes Office qui prennent en charge l’ensemble de conditions requises OneNoteApi, version 1.1.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Séries de conditions requises des API communes pour Office

La plupart des fonctionnalités du complément PowerPoint proviennent de la série courante d’API. Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour PowerPoint](/javascript/api/powerpoint)
- [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Spécification des exigences en matière d’hôtes Office et d’API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifeste XML des compléments Office](/office/dev/add-ins/develop/add-in-manifests)
