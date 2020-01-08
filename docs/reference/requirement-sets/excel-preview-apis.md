---
title: Version d’évaluation API JavaScript Excel
description: Informations détaillées sur les API JavaScript pour Excel à venir
ms.date: 01/02/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 7b751cfa550813fe3cb57f0d90d40e5632baf953
ms.sourcegitcommit: abe8188684b55710261c69e206de83d3a6bd2ed3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/08/2020
ms.locfileid: "40969544"
---
# <a name="excel-javascript-preview-apis"></a>Version d’évaluation API JavaScript Excel

Les nouvelles Excel JavaScript APIs introduits dans « Aperçu » et versions ultérieures deviennent partie intégrante d’une exigence spécifique, numérotée définie après que test suffisantes se produit et les commentaires des utilisateurs sont acquis.

Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Paramètres de culture | Obtient les paramètres du système culturel pour le classeur, tels que la mise en forme des nombres. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [](/javascript/api/excel/excel.numberformatinfo) [application](/javascript/api/excel/excel.application) NumberFormatInfo |
| [Insérer un classeur](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Insérer un classeur dans un autre.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| Classeur [enregistrer](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview) et [fermer](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) | Enregistrez et fermez ses classeurs.  | [Workbook](/javascript/api/excel/excel.workbook) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API JavaScript pour Excel actuellement en version préliminaire. Pour afficher la liste complète de toutes les API JavaScript pour Excel (y compris les API d’aperçu et les API précédemment publiées), voir [toutes les API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview).

| Class | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Fournit des informations basées sur les paramètres de culture système actuels. Cela inclut les noms de culture, la mise en forme de numéros et d’autres paramètres dépendants de la culture.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|Obtient la chaîne utilisée comme séparateur décimal pour les valeurs numériques. Cette fonction est basée sur les paramètres locaux d’Excel.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|Obtient la chaîne utilisée pour séparer les groupes de chiffres à gauche du séparateur décimal pour les valeurs numériques. Cette fonction est basée sur les paramètres locaux d’Excel.|
||[UseSystemSeparators,](/javascript/api/excel/excel.application#usesystemseparators)|Indique si les séparateurs système de Microsoft Excel sont activés.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (dimension : Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Obtient les valeurs d’une dimension unique de la série de graphiques. Il peut s’agir de valeurs de catégorie ou de valeurs de données, en fonction de la dimension spécifiée et de la façon dont les données sont mappées pour la série de graphiques.|
|[Comment](/javascript/api/excel/excel.comment)|[évaluation](/javascript/api/excel/excel.comment#resolved)|Obtient ou définit l’état du thread de commentaire. La valeur « true » signifie que le thread de commentaire est dans l’État résolu.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[évaluation](/javascript/api/excel/excel.commentreply#resolved)|Obtient ou définit l’état de la réponse de commentaire. La valeur « true » signifie que la réponse au commentaire est dans l’État résolu.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Obtient le nom de la culture au format languagecode2-Country/regioncode2 (par exemple « zh-CN » ou « en-US »). Cette fonction est basée sur les paramètres système actuels.|
||[numberFormatInfo](/javascript/api/excel/excel.cultureinfo#numberformatinfo)|Définit le format d’affichage des nombres approprié pour la culture. Cette fonction est basée sur les paramètres de culture actuelle du système.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Obtient la chaîne utilisée comme séparateur décimal pour les valeurs numériques. Cette fonction est basée sur les paramètres système actuels.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Obtient la chaîne utilisée pour séparer les groupes de chiffres à gauche du séparateur décimal pour les valeurs numériques. Cette fonction est basée sur les paramètres système actuels.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtient une cellule unique dans le tableau croisé dynamique basé sur une hiérarchie de données ainsi que les éléments de ligne et de colonne de leurs hiérarchies respectives. La cellule renvoyée est l’intersection de la ligne donnée et une colonne qui contient les données à partir de la hiérarchie donnée. Cette méthode est l’inverse de l’appel getPivotItems et getDataHierarchy sur une cellule particulière.|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtient l’objet de la plage contenant la cellule d’ancrage d’une cellule prise renversée dans. Échoue si appliqué à une plage comportant plusieurs cellules. En lecture seule.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Obtient l’objet de la plage contenant la cellule d’ancrage d’une cellule prise renversée dans. En lecture seule.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage. Échoue si appliqué à une plage comportant plusieurs cellules. En lecture seule.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage. En lecture seule.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Représente si toutes les cellules ont une bordure renversée.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Représente si toutes les cellules sont enregistrées sous la forme d’une formule matricielle.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Crée un graphique de fichiers SVG (SVG) à partir d’une chaîne XML et il est ajouté à la feuille de calcul. Renvoie un objet Forme qui représente la nouvelle image.|
|[Segment](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Représente le nom du segment utilisé dans la formule.|
|[Tableau](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Modifie le tableau pour utiliser le style de tableau par défaut.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Se produit lorsque le filtre est appliqué sur une table spécifique.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Se produit lorsque le filtre est appliqué sur n’importe quel tableau dans un classeur ou une feuille de calcul.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Représente l’ID de la table dans laquelle le filtre est appliqué.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Représente le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Représente l’id de la feuille de calcul qui contient le tableau.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Fermer (closeBehavior ? : Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fermer le classeur actif.|
||[Enregistrer (saveBehavior ? : Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Enregistrer le classeur actif.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True si le classeur utilise le calendrier depuis 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Obtient une collection de propriétés personnalisées au niveau de la feuille de calcul.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Se produit lorsque le filtre est appliqué sur un tableau spécifique.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Survient lorsque l’état masqué d’une ou plusieurs lignes a été modifié sur une feuille de calcul spécifique.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[adresse](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|Adresse de la plage qui a terminé le calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Se produit lorsqu’un filtre de la feuille de calcul est appliqué dans le classeur.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Survient lorsque l’état masqué d’une ou plusieurs lignes a été modifié sur une feuille de calcul spécifique.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Obtient la clé de la propriété personnalisée. En lecture seule.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Obtient la valeur de la propriété personnalisée. En lecture seule.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Obtient le nombre de propriétés personnalisées sur cette feuille de calcul.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse. Lève une exception si la propriété personnalisée n’existe pas.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse. Renvoie un objet null si la propriété personnalisée n’existe pas.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Représente le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Indique le nom du tableau auquel le filtre est appliqué.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Obtient le type de modification qui représente la manière dont l’événement a été déclenché. Pour `Excel.RowHiddenChangeType` plus d’informations, voir.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)
