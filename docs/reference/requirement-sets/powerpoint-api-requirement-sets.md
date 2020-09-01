---
title: Séries de conditions requises de l’API JavaScript pour PowerPoint
description: En savoir plus sur les ensembles de conditions requises de l’API JavaScript pour PowerPoint.
ms.date: 07/10/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: b2b5d4b7b5a0677812f227b6a32683c35bbf1662
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293505"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a>Séries de conditions requises de l’API JavaScript pour PowerPoint

Les ensembles de conditions requises sont des groupes nommés de membres de l’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’un complément nécessite. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Le tableau suivant répertorie les ensembles de conditions requises pour PowerPoint, les applications clientes Office qui prennent en charge ces ensembles de conditions requises et les numéros de version ou la date de disponibilité.

|  Ensemble de conditions requises  |  Office pour Windows<br>(connecté à un abonnement Microsoft 365)  |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac<br>(connecté à un abonnement Microsoft 365)  | Office sur le web |
|:-----|-----|:-----|:-----|:-----|:-----|
| PowerPointApi 1.1 | Version 1810 (Build 11001.20074) ou version ultérieure | 2.17 ou version ultérieure | 16.19 ou version ultérieure | Octobre 2018 |

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour plus d’informations sur les versions et les numéros de build d’Office, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a>API JavaScript pour PowerPoint 1.1

L’API JavaScript PowerPoint 1.1 inclut une seule API pour créer une nouvelle présentation. Pour plus de détails sur l’API, voir [API JavaScript pour PowerPoint](../../powerpoint/powerpoint-add-ins.md).

## <a name="runtime-requirement-support-check"></a>Vérification de la prise en charge d’une exigence d'exécution

Lors de l’exécution, les compléments peuvent vérifier si une application spécifique prend en charge un ensemble de conditions requises de l’API en procédant comme suit.

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>Vérification de la prise en charge des conditions requises basée sur le manifeste

Utilisez l’élément `Requirements` dans le manifeste du complément pour spécifier des ensembles de conditions requises essentiels ou des membres d’API que votre complément doit utiliser. Si l’application ou la plateforme Office ne prend pas en charge les ensembles de conditions requises ou les membres d’API spécifiés dans l’élément `Requirements`, le complément ne s’exécute pas dans cette application ou plateforme et ne s’affiche pas dans Mes compléments.

Cet exemple de code illustre un complément qui se charge dans toutes les applications clientes Office qui prennent en charge l’ensemble de conditions requises OneNoteApi, version 1.1.

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
- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécifier les exigences en matière d’applications Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
