---
title: Version d’évaluation API JavaScript Excel
description: Informations détaillées sur les API JavaScript pour Excel à venir.
ms.date: 10/26/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a1cb3afb28f69ff5b0c0bd03bfae9877dda91906
ms.sourcegitcommit: a4e09546fd59579439025aca9cc58474b5ae7676
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/27/2020
ms.locfileid: "48774739"
---
# <a name="excel-javascript-preview-apis"></a>Version d’évaluation API JavaScript Excel

Les nouvelles Excel JavaScript APIs introduits dans « Aperçu » et versions ultérieures deviennent partie intégrante d’une exigence spécifique, numérotée définie après que test suffisantes se produit et les commentaires des utilisateurs sont acquis.

Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Types de données liées | Prend en charge les types de données connectés à Excel à partir de sources externes. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Affichages de feuille nommée | Fournit un contrôle par programme des affichages de feuille de calcul par utilisateur. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API JavaScript pour Excel actuellement en version préliminaire. Pour obtenir la liste complète des API JavaScript pour Excel (dont les API d’aperçu et les API précédemment publiées), voir [toutes les API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true).

| Class | Champs | Description |
|:---|:---|:---|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|Nom du fournisseur de données pour le type de données liées. Cela peut changer lorsque les informations sont récupérées à partir du service.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|Date et heure locales du fuseau horaire depuis l’ouverture du classeur lors de la dernière actualisation du type de données liées.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|Nom du type de données liées. Cela peut changer lorsque les informations sont récupérées à partir du service.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|Fréquence, en secondes, à laquelle le type de données liées est actualisé si `refreshMode` est défini sur « périodique ».|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|Mécanisme par lequel les données du type de données liées sont récupérées.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|ID unique du type de données liées.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Renvoie un tableau avec tous les modes d’actualisation pris en charge par le type de données liées. Le contenu du tableau peut changer lorsque les informations sont récupérées à partir du service.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Effectue une demande pour actualiser le type de données liées. Si le service est occupé ou inaccessible temporairement, la demande ne sera pas remplie.|
||[requestSetRefreshMode (refreshMode : Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Effectue une demande pour modifier le mode d’actualisation de ce type de données liées.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|ID unique du nouveau type de données liées.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Obtient le nombre de types de données liées dans la collection.|
||[getItem (Key : nombre)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Obtient un type de données liées par ID de service.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Obtient un type de données liées par son index dans la collection.|
||[getItemOrNullObject (Key : nombre)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Obtient un type de données liées par ID. Si le type de données liées n’existe pas, il s’agit d’un objet dont la `isNullObject` propriété a la valeur `true` . Pour plus d’informations, consultez la rubrique {@link https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties | * Méthodes et propriétés de OrNullObject}.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Effectue une demande d’actualisation de tous les types de données liées dans la collection.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Active l’affichage tableau. Cela équivaut à l’utilisation de l’option « Basculer vers » dans l’interface utilisateur Excel.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Supprime l’affichage tableau de la feuille de calcul.|
||[doublon (Name ?: String)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Crée une copie de l’affichage de cette feuille.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Obtient ou définit le nom de l’affichage tableau.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Crée une nouvelle vue de feuille portant le nom donné.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Crée et active un nouvel affichage de tableau temporaire.|
||[Exit ()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Quitte l’affichage de la feuille active.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Obtient l’affichage de la feuille actuellement actif de la feuille de calcul.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Obtient le nombre d’affichages de feuille dans cette feuille de calcul.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Obtient un affichage tableau à l’aide de son nom.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Obtient un affichage feuille par son index dans la collection.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|Description du texte de remplacement du tableau croisé dynamique.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|Titre de texte de remplacement du tableau croisé dynamique.|
||[displayBlankLineAfterEachItem (Display : Boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Définit si une ligne vide doit être affichée après chaque élément. Cette valeur est définie au niveau global pour le tableau croisé dynamique et appliquée à des champs PivotFields individuels.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|Texte qui est rempli automatiquement dans une cellule vide du tableau croisé dynamique si `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Indique si les cellules vides dans le tableau croisé dynamique doivent être renseignées avec le `emptyCellText` . Elle a la valeur False par défaut.|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtient une cellule unique dans le tableau croisé dynamique basé sur une hiérarchie de données ainsi que les éléments de ligne et de colonne de leurs hiérarchies respectives. La cellule renvoyée est l’intersection de la ligne donnée et une colonne qui contient les données à partir de la hiérarchie donnée. Cette méthode est l’inverse de l’appel getPivotItems et getDataHierarchy sur une cellule particulière.|
||[repeatAllItemLabels (repeatLabels : booléen)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Définit le paramètre « répéter toutes les étiquettes d’éléments » sur tous les champs du tableau croisé dynamique.|
||[setStyle (style : String \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Définit le style appliqué au tableau croisé dynamique.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Indique si le tableau croisé dynamique affiche les en-têtes de champ (légendes de champ et listes déroulantes de filtre).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Indique si le tableau croisé dynamique est actualisé lors de l’ouverture du classeur. Correspond au paramètre « actualiser lors du chargement » dans l’interface utilisateur.|
|[Range](/javascript/api/excel/excel.range)|[getMergedAreas()](/javascript/api/excel/excel.range#getmergedareas--)|Renvoie un `RangeAreas` objet qui représente les zones fusionnées dans cette plage. Notez que si le nombre de zones fusionnées dans cette plage est supérieur à 512, l’API ne renverra pas le résultat.|
||[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|Renvoie un `WorkbookRangeAreas` Object qui représente la plage contenant tous les antécédents d’une cellule dans une même feuille de calcul ou dans plusieurs feuilles de calcul.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|Mode d’actualisation du type de données liées.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|ID unique de l’objet dont le mode d’actualisation a été modifié.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[Actualisé](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indique si la demande d’actualisation a réussi.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|ID unique de l’objet dont la demande d’actualisation a été exécutée.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[affichés](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Tableau qui contient tous les avertissements générés à partir de la demande d’actualisation.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Crée un graphique de fichiers SVG (SVG) à partir d’une chaîne XML et il est ajouté à la feuille de calcul. Renvoie un objet Forme qui représente la nouvelle image.|
|[Segment](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Représente le nom du segment utilisé dans la formule.|
||[setStyle (style : String \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Définit le style appliqué au segment.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Modifie le tableau pour utiliser le style de tableau par défaut.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Se produit lorsque le filtre est appliqué sur une table spécifique.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Style appliqué au tableau.|
||[setStyle (style : String \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Définit le style appliqué au tableau.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Se produit lorsque le filtre est appliqué sur n’importe quel tableau dans un classeur ou une feuille de calcul.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Obtient l’ID de la table dans laquelle le filtre est appliqué.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Obtient l’ID de la feuille de calcul qui contient le tableau.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Renvoie une collection de types de données liées qui font partie du classeur.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Indique si le volet de liste de champs du tableau croisé dynamique est affiché au niveau du classeur.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True si le classeur utilise le calendrier depuis 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Renvoie une collection de vues de feuille présentes dans la feuille de calcul.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Se produit lorsque le filtre est appliqué sur un tableau spécifique.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Se produit lorsqu’un filtre de la feuille de calcul est appliqué dans le classeur.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle le filtre est appliqué.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
