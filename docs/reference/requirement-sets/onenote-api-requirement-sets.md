---
title: Ensembles de conditions requises de l’API JavaScript pour OneNote
description: En savoir plus sur les ensembles de conditions requises de l’API JavaScript pour OneNote.
ms.date: 08/24/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: c8cadacac640cbe710c9894a65ee780267066afc
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293526"
---
# <a name="onenote-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour OneNote

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’ils nécessitent. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Le tableau suivant répertorie les ensembles de conditions requises pour OneNote, les applications clientes Office qui prennent en charge ces conditions ainsi que les numéros de version ou la date de disponibilité.

|  Ensemble de conditions requises  |  Office sur le web |
|:-----|:-----|
| [OneNoteApi 1.1](/javascript/api/onenote?view=onenote-js-1.1)  | Septembre 2016 |  

## <a name="onenote-javascript-api-11"></a>API JavaScript pour OneNote 1.1

L’API JavaScript 1.1 pour OneNote est la première version de l’API. Pour plus d’informations sur l’API, consultez les rubriques de référence sur l’[Récapitulatif de programmation API JavaScript pour OneNote](../../onenote/onenote-add-ins-programming-overview.md).

## <a name="runtime-requirement-support-check"></a>Vérification de la prise en charge d’une exigence d'exécution

Lors de l’exécution, les compléments peuvent vérifier si une application Office spécifique prend en charge une série de conditions requises d’API en procédant comme suit.

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
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
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
