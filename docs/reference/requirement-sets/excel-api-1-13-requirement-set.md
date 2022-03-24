---
title: Excel l’ensemble de conditions requises de l’API JavaScript 1.13
description: Détails sur l’ensemble de conditions requises ExcelApi 1.13.
ms.date: 07/09/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 5d7358c35dc4560bf5478bb9ad9970fc364a1b6a
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63747045"
---
# <a name="whats-new-in-excel-javascript-api-113"></a>Nouveautés de l Excel API JavaScript 1.13

ExcelApi 1.13 a ajouté une méthode pour insérer des feuilles de calcul dans un workbook à partir d’une chaîne encodée en Base64 et un événement pour détecter l’activation du workbook. Il a également augmenté la prise en charge des formules dans les plages en ajoutant des API pour suivre les modifications apportées aux formules et localiser les cellules dépendantes directes d’une formule. En outre, il a étendu la prise en charge du tableau croisé dynamique en ajoutant des API PivotLayout pour le texte de alt, le style et la gestion des cellules vides.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Événements de changement de formule](../../excel/excel-add-ins-worksheets.md#detect-formula-changes) | Suivre les modifications apportées aux formules, y compris la source et le type d’événement à l’origine d’une modification. | [Worksheet.onFormulaChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member)|
| [Dépendants des formules](../../excel/excel-add-ins-ranges-precedents-dependents.md#get-the-direct-dependents-of-a-formula) | Recherchez les cellules dépendantes directes d’une formule. | [Range.getDirectDependents](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1)) |
| [Insérer des feuilles de calcul](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one) | Insérez des feuilles de calcul à partir d’un autre workbook dans le workbook actuel sous la forme d’une chaîne codée en Base64. | [Workbook.insertWorksheetsFromBase64](/javascript/api/excel/excel.workbook#excel-excel-workbook-insertworksheetsfrombase64-member(1)) |
| [PivotLayout de tableau croisé dynamique](../../excel/excel-add-ins-pivottables.md#other-pivotlayout-functions) | Développement de la classe PivotLayout, y compris la nouvelle prise en charge du texte de alt et de la gestion des cellules vides. | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de Excel l’ensemble de conditions requises de l’API JavaScript 1.13. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.13 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.13](/javascript/api/excel?view=excel-js-1.13&preserve-view=true) ou une version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#excel-excel-formulachangedeventdetail-celladdress-member)|Adresse de la cellule qui contient la formule modifiée.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#excel-excel-formulachangedeventdetail-previousformula-member)|Représente la formule précédente, avant qu’elle n’a été modifiée.|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#excel-excel-insertworksheetoptions-positiontype-member)|Position d’insertion, dans le livre de calcul actuel, des nouvelles feuilles de calcul.|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#excel-excel-insertworksheetoptions-relativeto-member)|Feuille de calcul du manuel actuel référencé pour le `WorksheetPositionType` paramètre.|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#excel-excel-insertworksheetoptions-sheetnamestoinsert-member)|Noms des feuilles de calcul individuelles à insérer.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-alttextdescription-member)|Description de texte de alt du tableau croisé dynamique.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-alttexttitle-member)|Titre de texte de alt du tableau croisé dynamique.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-displayblanklineaftereachitem-member(1))|Définit si une ligne vide doit être affichée après chaque élément.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-emptycelltext-member)|Texte qui est automatiquement rempli dans n’importe quelle cellule vide du tableau croisé dynamique si `fillEmptyCells == true`.|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-fillemptycells-member)|Spécifie si les cellules vides du tableau croisé dynamique doivent être remplies avec le `emptyCellText`.|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-repeatallitemlabels-member(1))|Définit le paramètre « Répéter toutes les étiquettes d’éléments » dans tous les champs du tableau croisé dynamique.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showfieldheaders-member)|Spécifie si le tableau croisé dynamique affiche les en-têtes de champ (légendes de champ et les drop-downs de filtre).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-refreshonopen-member)|Spécifie si le tableau croisé dynamique est actualisé à l’ouverture du manuel.|
|[Range](/javascript/api/excel/excel.range)|[getDirectDependents()](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1))|Renvoie un `WorkbookRangeAreas` objet qui représente la plage contenant tous les dépendants directs d’une cellule dans la même feuille de calcul ou dans plusieurs feuilles de calcul.|
||[getExtendedRange(direction: Excel. KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getextendedrange-member(1))|Renvoie un objet de plage qui inclut la plage actuelle et jusqu’au bord de la plage, en fonction de la direction fournie.|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#excel-excel-range-getmergedareasornullobject-member(1))|Renvoie un objet RangeAreas qui représente les zones fusionnées dans cette plage.|
||[getRangeEdge(direction: Excel. KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getrangeedge-member(1))|Renvoie un objet de plage qui est la cellule edge de la zone de données qui correspond au sens fourni.|
|[Table](/javascript/api/excel/excel.table)|[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#excel-excel-table-resize-member(1))|Resize the table to the new range.|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: chaîne, options ? : Excel. InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#excel-excel-workbook-insertworksheetsfrombase64-member(1))|Insère les feuilles de calcul spécifiées à partir d’un workbook source dans le workbook actuel.|
||[onActivated](/javascript/api/excel/excel.workbook#excel-excel-workbook-onactivated-member)|Se produit lorsque le workbook est activé.|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#excel-excel-workbookactivatedeventargs-type-member)|Obtient le type de l’événement.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFormulaChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member)|Se produit lorsqu’une ou plusieurs formules sont modifiées dans cette feuille de calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformulachanged-member)|Se produit lorsqu’une ou plusieurs formules sont modifiées dans une feuille de calcul de cette collection.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-formuladetails-member)|Obtient un tableau d’objets `FormulaChangedEventDetail` , qui contient les détails sur toutes les formules modifiées.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-source-member)|Source de l'événement.|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul dans laquelle la formule a été modifiée.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
