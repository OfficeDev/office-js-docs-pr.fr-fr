---
title: Séries de conditions requises de l’API JavaScript pour PowerPoint
description: En savoir plus sur les ensembles de conditions requises de l’API JavaScript pour PowerPoint.
ms.date: 10/26/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: cf9ab510e4b35a140c77ee958279cb85a2189fa2
ms.sourcegitcommit: a4e09546fd59579439025aca9cc58474b5ae7676
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/27/2020
ms.locfileid: "48774727"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a>Séries de conditions requises de l’API JavaScript pour PowerPoint

Les ensembles de conditions requises sont des groupes nommés de membres de l’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’un complément nécessite. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Le tableau suivant répertorie les ensembles de conditions requises pour PowerPoint, les applications clientes Office qui prennent en charge ces ensembles de conditions requises et les numéros de version ou la date de disponibilité.

|  Ensemble de conditions requises  |  Office pour Windows<br>(connecté à un abonnement Microsoft 365)  |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac<br>(connecté à un abonnement Microsoft 365)  | Office sur le web |
|:-----|-----|:-----|:-----|:-----|:-----|
| [Aperçu](powerpoint-preview-apis.md)  | Veuillez utiliser la dernière version d’Office pour tester les API d’aperçu (vous devrez peut-être adhérer au [programme Office Insider](https://insider.office.com)). |
| PowerPointApi 1.1 | Version 1810 (Build 11001.20074) ou version ultérieure | 2.17 ou version ultérieure | 16.19 ou version ultérieure | Octobre 2018 |

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour plus d’informations sur les versions et les numéros de build d’Office, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a>API JavaScript pour PowerPoint 1.1

L’API JavaScript PowerPoint 1.1 inclut une [seule API pour créer une nouvelle présentation](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-). Pour plus d’informations sur l’API, consultez [Créer une présentation](../../powerpoint/powerpoint-add-ins.md#create-a-presentation).

## <a name="how-to-use-powerpoint-requirement-sets-at-runtime-and-in-the-manifest"></a>Utiliser les conditions requises PowerPoint au moment de l’exécution et dans le manifeste

> [!NOTE]
> Cette section suppose que vous êtes familiarisé avec les rubriques [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md) et [Spécifier les applications Office et les exigences de l’API](../../develop/specify-office-hosts-and-api-requirements.md).

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Le complément Office peut effectuer une vérification à l’exécution ou utiliser des ensembles de conditions requises spécifiés dans le manifeste pour déterminer si une application Office prend en charge les API requises par le complément.

### <a name="checking-for-requirement-set-support-at-runtime"></a>Vérification de la prise en charge de l’ensemble de conditions requises à l’exécution

L’exemple de code suivant montre comment déterminer si l’application Office dans laquelle le complément est en cours d’exécution prend en charge l’ensemble spécifié de conditions requises pour l’API.

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
} else {
  // Provide alternate flow/logic.
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>Définition de la prise en charge de l’ensemble de conditions requises dans le manifeste

Vous pouvez utiliser l’[élément Requirements](../manifest/requirements.md) dans le manifeste de complément pour spécifier les ensembles de conditions requises minimales et/ou les méthodes d’API que votre complément doit activer. Si la plateforme ou l’application Office ne prend pas en charge les ensembles de conditions requises ou les méthodes d’API spécifiées dans l’élément `Requirements` du manifeste, le complément ne s’exécute pas dans cette application ou plateforme et ne s’affiche pas dans la liste de compléments dans **Mes compléments** .Si votre complément requiert une configuration spécifique pour les fonctionnalités complètes, mais qu’il peut fournir une valeur même pour les utilisateurs sur les plateformes qui ne prennent pas en charge la condition requise, nous vous recommandons de vérifier la prise en charge des exigences au moment de l’exécution, comme décrit ci-dessus, au lieu de définir la prise en charge de la condition requise dans le manifeste.

L’exemple de code suivant montre l’élément `Requirements` dans un manifeste indiquant que le complément doit être chargé dans toutes les applications clientes Office prenant en charge l’ensemble de conditions requises PowerPointApi version 1.1 ou ultérieure.

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
- [Spécifier les applications Office et les exigences de l’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
