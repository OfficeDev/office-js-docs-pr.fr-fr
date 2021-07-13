---
title: Ensembles de conditions requises de l’API JavaScript pour OneNote
description: En savoir plus sur les ensembles de conditions requises de l’API JavaScript pour OneNote.
ms.date: 08/24/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: ecdb26edca54758540688ba03b1d9c1eec14e739
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350189"
---
# <a name="onenote-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour OneNote

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’ils nécessitent. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Le tableau suivant répertorie les ensembles de conditions requises pour OneNote, les applications clientes Office qui prennent en charge ces conditions ainsi que les numéros de version ou la date de disponibilité.

|  Ensemble de conditions requises  |  Office sur le web |
|:-----|:-----|
| [OneNoteApi 1.1](/javascript/api/onenote?view=onenote-js-1.1&preserve-view=true)  | Septembre 2016 |  

## <a name="onenote-javascript-api-11"></a>API JavaScript pour OneNote 1.1

L’API JavaScript 1.1 pour OneNote est la première version de l’API. Pour plus de détails sur l’API, consultez les [vue d’ensemble de l’API JavaScript pour OneNote](../../onenote/onenote-add-ins-programming-overview.md).

## <a name="runtime-requirement-support-check"></a>Vérification de la prise en charge d’un ensemble de conditions requises à l’exécution

Lors de l’exécution, les compléments peuvent vérifier si une application Office spécifique prend en charge un ensemble de conditions requises de l’API en procédant comme suit:

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>Vérification de la prise en charge des conditions requises basée sur le manifeste

Utilisez `Requirements`l’élément dans le manifeste du complément pour spécifier des ensembles de conditions requises critiques ou des membres de l’API que votre complément doit utiliser. Si l’application Office ou la plateforme ne prend pas en charge les ensembles de conditions requises ou les membres de l’API spécifiés dans`Requirements` l’élément, le complément ne s’exécutera pas dans cet application ou cette plateforme et ne s’affichera pas dans mes compléments.

Cet exemple de code illustre un complément qui se charge dans toutes les applications clientes Office qui prennent en charge l’ensemble de conditions requises OneNoteApi, version 1.1.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Séries de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour OneNote](/javascript/api/onenote)
- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécifier les exigences en matière d’applications Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
