---
title: Excel Ensemble de conditions requises de l’API JavaScript 1.13
description: Détails sur l’ensemble de conditions requises ExcelApi 1.13.
ms.date: 07/09/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 422cc8da19ac901de68cdfa59d7ab9670858de6f
ms.sourcegitcommit: 95fc1fc8a0dbe8fc94f0ea647836b51cc7f8601d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/14/2021
ms.locfileid: "53418698"
---
# <a name="whats-new-in-excel-javascript-api-113"></a>Nouveautés de l Excel API JavaScript 1.13

ExcelApi 1.13 a ajouté une méthode pour insérer des feuilles de calcul dans un workbook à partir d’une chaîne encodée en Base64 et un événement pour détecter l’activation du workbook. Il a également augmenté la prise en charge des formules dans les plages en ajoutant des API pour suivre les modifications apportées aux formules et localiser les cellules dépendantes directes d’une formule. En outre, il a étendu la prise en charge du tableau croisé dynamique en ajoutant des API PivotLayout pour le texte de alt, le style et la gestion des cellules vides.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Événements de changement de formule](../../excel/excel-add-ins-worksheets.md#detect-formula-changes) | Suivre les modifications apportées aux formules, y compris la source et le type d’événement à l’origine d’une modification. | [Worksheet.onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged)|
| [Dépendants des formules](../../excel/excel-add-ins-ranges-precedents-dependents.md#get-the-direct-dependents-of-a-formula) | Recherchez les cellules dépendantes directes d’une formule. | [Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__) |
| [Insérer des feuilles de calcul](../../excel//excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one) | Insérez des feuilles de calcul à partir d’un autre workbook dans le workbook actuel sous la forme d’une chaîne codée en Base64. | [Workbook.insertWorksheetsFromBase64](/javascript/api/excel/excel.workbook#insertWorksheetsFromBase64_base64File__options_) |
| [PivotLayout de tableau croisé dynamique](../../excel/excel-add-ins-pivottables.md#other-pivotlayout-functions) | Développement de la classe PivotLayout, y compris la nouvelle prise en charge du texte de alt et de la gestion des cellules vides. | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de Excel l’ensemble de conditions requises de l’API JavaScript 1.13. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.13 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.13](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)ou version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|Adresse de la cellule qui contient la formule modifiée.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|Représente la formule précédente, avant qu’elle n’a été modifiée.|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#positiontype)|Position d’insertion, dans le livre de calcul actuel, des nouvelles feuilles de calcul.|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#relativeto)|Feuille de calcul du manuel actuel référencé pour le `WorksheetPositionType` paramètre.|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetnamestoinsert)|Noms des feuilles de calcul individuelles à insérer.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|Description de texte de alt du tableau croisé dynamique.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|Titre de texte de alt du tableau croisé dynamique.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Définit si une ligne vide doit être affichée après chaque élément.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|Texte qui est automatiquement rempli dans n’importe quelle cellule vide du tableau croisé dynamique si `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Spécifie si les cellules vides du tableau croisé dynamique doivent être remplies avec le `emptyCellText` .|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Définit le paramètre « Répéter toutes les étiquettes d’éléments » sur tous les champs du tableau croisé dynamique.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Spécifie si le tableau croisé dynamique affiche les en-têtes de champ (légendes de champ et les drop-downs de filtre).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Spécifie si le tableau croisé dynamique est actualisé à l’ouverture du manuel.|
|[Range](/javascript/api/excel/excel.range)|[getDirectDependents()](/javascript/api/excel/excel.range#getdirectdependents--)|Renvoie un objet qui représente la plage contenant tous les dépendants directs d’une cellule dans la même feuille de calcul ou `WorkbookRangeAreas` dans plusieurs feuilles de calcul.|
||[getExtendedRange(direction: Excel. KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|Renvoie un objet de plage qui inclut la plage actuelle et jusqu’au bord de la plage, en fonction de la direction fournie.|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#getmergedareasornullobject--)|Renvoie un objet RangeAreas qui représente les zones fusionnées dans cette plage.|
||[getRangeEdge(direction: Excel. KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|Renvoie un objet de plage qui est la cellule edge de la zone de données qui correspond à la direction fournie.|
|[Table](/javascript/api/excel/excel.table)|[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#resize-newrange-)|Resize the table to the new range.|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel. InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#insertworksheetsfrombase64-base64file--options-)|Insère les feuilles de calcul spécifiées à partir d’un workbook source dans le workbook actuel.|
||[onActivated](/javascript/api/excel/excel.workbook#onactivated)|Se produit lorsque le workbook est activé.|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#type)|Obtient le type de l’événement.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|Se produit lorsqu’une ou plusieurs formules sont modifiées dans cette feuille de calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|Se produit lorsqu’une ou plusieurs formules sont modifiées dans une feuille de calcul de cette collection.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|Obtient un tableau `FormulaChangedEventDetail` d’objets, qui contient les détails sur toutes les formules modifiées.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|Source de l'événement.|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle la formule a été modifiée.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
