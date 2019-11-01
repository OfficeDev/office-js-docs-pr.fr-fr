---
title: Problèmes de codage courants et comportements de plateforme inattendus
description: Liste des problèmes de plateforme d’API JavaScript pour Office fréquemment rencontrés par les développeurs.
ms.date: 10/29/2019
localization_priority: Normal
ms.openlocfilehash: 8cea95e3214585ba8e0b77535916f9c564dde9df
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902152"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a>Problèmes de codage courants et comportements de plateforme inattendus

Cet article met en évidence les aspects de l’API JavaScript pour Office qui peuvent entraîner un comportement inattendu ou nécessiter des modèles de codage spécifiques pour obtenir le résultat souhaité. Si vous rencontrez un problème qui se trouve dans cette liste, faites-le nous connaître en utilisant le formulaire de commentaires au bas de l’article.

## <a name="some-properties-must-be-set-with-json-structs"></a>Certaines propriétés doivent être définies avec des structs JSON

> [!NOTE]
> Cette section s’applique uniquement aux API propres à l’hôte pour Excel et Word.

Certaines propriétés doivent être définies en tant que structs JSON, au lieu de définir leurs sous-propriétés individuelles. Vous trouverez un exemple dans [PageLayout](/javascript/api/excel/excel.pagelayout). La `zoom` propriété doit être définie avec un seul objet [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , comme illustré ci-dessous :

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

Dans l’exemple précédent, vous ne seriez ***pas*** en mesure d' `zoom` affecter directement une `sheet.pageLayout.zoom.scale = 200;`valeur :. Cette instruction génère une erreur car `zoom` elle n’est pas chargée. Même si `zoom` elles ont été chargées, l’ensemble de l’étendue ne prendra pas effet. Toutes les opérations de `zoom`contexte se produisent, actualisant l’objet proxy dans le complément et remplaçant les valeurs définies localement.

Ce comportement diffère des [Propriétés de navigation](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) telles que [Range. format](/javascript/api/excel/excel.range#format). Les propriétés `format` de peuvent être définies à l’aide de la navigation d’objet, comme illustré ci-dessous :

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

Vous pouvez identifier une propriété dont les propriétés subordonnées doivent être définies avec un struct JSON en vérifiant son modificateur en lecture seule. Les propriétés non en lecture seule de toutes les propriétés en lecture seule peuvent être définies directement. Les propriétés accessibles en `PageLayout.zoom` écriture comme doivent être définies avec une structure JSON. En Résumé :

- Propriété en lecture seule : les sous-propriétés peuvent être définies via la navigation.
- Propriété accessible en écriture : les sous-propriétés doivent être définies avec une structure JSON (et ne peuvent pas être définies via la navigation).

## <a name="setting-read-only-properties"></a>Définition de propriétés en lecture seule

Les [définitions](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) de la machine à écrire pour Office js spécifient les propriétés d’objet en lecture seule. Si vous tentez de définir une propriété en lecture seule, l’opération d’écriture échoue sans avertissement, sans qu’aucune erreur ne soit générée. L’exemple suivant tente à tort de définir la propriété en lecture seule [Chart.ID](/javascript/api/excel/excel.chart#id).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a>Voir aussi

- [OfficeDev/Office-js](https://github.com/OfficeDev/office-js/issues): le lieu de signaler et d’afficher les problèmes liés à la plateforme des compléments Office et aux API JavaScript.
- [Débordement de pile](https://stackoverflow.com/questions/tagged/office-js): emplacement où poser des questions de programmation sur les API JavaScript Office. Veillez à appliquer la balise « Office-js » à votre question lors de la publication dans le débordement de pile.
- [UserVoice](https://officespdev.uservoice.com/): le lieu de suggérer de nouvelles fonctionnalités pour la plateforme des compléments Office et les API JavaScript pour Office.
