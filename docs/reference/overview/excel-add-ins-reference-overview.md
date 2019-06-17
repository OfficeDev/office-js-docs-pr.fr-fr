---
title: Vue d’ensemble de l’API JavaScript pour Excel
description: ''
ms.date: 06/10/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: aa9574a93252c0011b211c39e37cc013beb64432
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910146"
---
# <a name="excel-javascript-api-overview"></a>Vue d’ensemble de l’API JavaScript pour Excel

Vous pouvez utiliser l’API JavaScript pour Excel pour créer des compléments pour Excel 2016 ou version ultérieure. La liste suivante affiche les objets de haut niveau Excel qui sont disponibles dans l’API. Chaque lien vers la page d’un objet contient une description des propriétés, des événements et des méthodes disponibles sur l’objet. Utilisez les liens dans le menu pour en savoir plus.

Certains objets Excel principaux sont répertoriés ci-après pour faciliter la tâche :

- [Workbook](/javascript/api/excel/excel.workbook) : objet de niveau supérieur qui contient les objets de classeur associés tels que les feuilles de calcul, les tableaux, les plages, etc. Il permet également d’établir la liste des références associées.

- [Worksheet](/javascript/api/excel/excel.worksheet) : représente une feuille de calcul dans un classeur.
  - [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) : collection des objets **Worksheet** dans un classeur.
  - [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) : représente la protection d’un objet **Worksheet**.

- [Range](/javascript/api/excel/excel.range) : représente une cellule, une ligne, une colonne ou une sélection de cellules contenant un ou plusieurs blocs contigus de cellules.
  - [ConditionalFormat](/javascript/api/excel/excel.conditionalformat) : objet définissant une règle et un format appliqués à la plage lorsque la condition de la règle est remplie.
  - [DataValidation](/javascript/api/excel/excel.datavalidation): objet qui limite l’intervention de l’utilisateur à une plage basée sur une série de critères.
  - [RangeSort](/javascript/api/excel/excel.rangesort) : représente un objet qui gère les opérations de tri sur une plage.

- [Table](/javascript/api/excel/excel.table) : représente une collection de cellules organisées conçue pour faciliter la gestion des données.
  - [TableCollection](/javascript/api/excel/excel.tablecollection) : collection de tableaux dans un classeur ou une feuille de calcul.
  - [TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection) : collection de toutes les colonnes d’un tableau.
  - [TableRowCollection](/javascript/api/excel/excel.tablerowcollection) : collection de toutes les lignes d’un tableau.
  - [TableSort](/javascript/api/excel/excel.tablesort) : représente un objet qui gère les opérations de tri sur un tableau.

- [Chart](/javascript/api/excel/excel.chart) : représente un objet graphique dans feuille de calcul, qui est une représentation visuelle de données sous-jacentes.
  - [ChartCollection](/javascript/api/excel/excel.chartcollection) : collection de graphiques d’une feuille de calcul.

- [PivotTable](/javascript/api/excel/excel.pivottable): représente un tableau croisé dynamique Excel, qui est un regroupement hiérarchique et une présentation de données.
  - [PivotTableCollection](/javascript/api/excel/excel.pivottablecollection) : collection de tableaux croisés dynamiques dans une feuille de calcul.

- [Filter](/javascript/api/excel/excel.filter) : représente un objet qui gère le filtrage de colonne d’un tableau.

- [NamedItem](/javascript/api/excel/excel.nameditem) : représente un nom défini pour une plage de cellules ou une valeur.
  - [NamedItemCollection](/javascript/api/excel/excel.nameditemcollection) : collection d’objets **NamedItem** dans un classeur.

- [Binding](/javascript/api/excel/excel.binding) : classe abstraite qui représente une liaison à une section du classeur.
  - [BindingCollection](/javascript/api/excel/excel.bindingcollection) : collection d’objets **Binding** dans un classeur.

## <a name="excel-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour Excel

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour en savoir plus sur les ensembles de conditions requises de l’API JavaScript pour Excel, consultez l’article [Ensembles de conditions requises de l’API JavaScript pour Excel](../requirement-sets/excel-api-requirement-sets.md).

## <a name="excel-javascript-api-reference"></a>Référence de l’API JavaScript pour Excel

Pour en savoir plus sur l’API JavaScript pour Excel, consultez la [documentation de référence de l’API JavaScript pour Excel](/javascript/api/excel).

## <a name="see-also"></a>Voir aussi

- [Présentation des compléments Excel](/office/dev/add-ins/excel/excel-add-ins-overview)
- [Vue d’ensemble de la plateforme des compléments Office](/office/dev/add-ins/overview/office-add-ins)
- [Exemples de compléments Excel sur GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
- [Spécifications ouvertes des API](../openspec/openspec.md)
