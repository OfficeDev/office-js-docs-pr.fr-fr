# <a name="excel-javascript-api-overview"></a>Vue d’ensemble de l’API JavaScript d’Excel

Vous pouvez utiliser l’API JavaScript d’Excel pour créer des compléments pour Excel 2016 ou version ultérieure. La liste suivante affiche les objets de haut niveau Excel qui sont disponibles dans l’API. Chaque page d’objet contient une description des propriétés, des événements et des méthodes disponibles sur l’objet. Utilisez les liens dans le menu pour en savoir plus.

Certains objets Excel principaux sont répertoriés ci-après pour faciliter la tâche : 

- [Workbook](/javascript/api/excel/excel.workbook) : objet de niveau supérieur qui contient les objets de classeur associés tels que les feuilles de calcul, les tableaux, les plages, etc. Il permet également d’établir la liste des références associées.

- [Worksheet](/javascript/api/excel/excel.worksheet) : représente une feuille de calcul dans un classeur. 
    - [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) : collection d’objets **Worksheet** dans un classeur.

- [Range](/javascript/api/excel/excel.range) : représente une cellule, une ligne, une colonne ou une sélection de cellules contenant des blocs contigus de cellules.

- [Table](/javascript/api/excel/excel.table) : représente une collection de cellules organisées conçue pour faciliter la gestion des données.
    - [TableCollection](/javascript/api/excel/excel.tablecollection) : collection de tableaux d’un classeur ou d’une feuille de calcul.
    - [TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection) : collection de toutes les colonnes d’un tableau.
    - [TableRowCollection](/javascript/api/excel/excel.tablerowcollection) : collection de toutes les lignes d’un tableau.

- [Chart](/javascript/api/excel/excel.chart) : représente un objet de graphique dans une feuille de calcul, qui est une représentation visuelle de données sous-jacentes.
    - [ChartCollection](/javascript/api/excel/excel.chartcollection) : collection de graphiques d’une feuille de calcul.

- [TableSort](/javascript/api/excel/excel.tablesort) : représente un objet qui gère les opérations de tri sur les objets **Table**.

- [RangeSort](/javascript/api/excel/excel.rangesort) : représente un objet qui gère les opérations de tri sur les objets **Range**.

- [Filter](/javascript/api/excel/excel.filter) : représente un objet qui gère le filtrage de colonne d’un tableau.

- [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) : représente la protection d’un objet **Worksheet**.

- [NamedItem](/javascript/api/excel/excel.nameditem) : représente un nom défini pour une plage de cellules ou une valeur. 
    - [NamedItemCollection](/javascript/api/excel/excel.nameditemcollection) : collection d’objets **NamedItem** dans un classeur.

- [Binding](/javascript/api/excel/excel.binding) : classe abstraite qui représente une liaison vers une section du classeur.
    - [BindingCollection](/javascript/api/excel/excel.bindingcollection) : collection d’objets **Binding** dans un classeur.

## <a name="excel-javascript-api-open-specifications"></a>Spécifications d’ouverture de l’API JavaScript pour Excel

Au fur et à mesure que nous concevons et développons de nouvelles API pour les compléments Excel, nous les mettons à votre disposition sur notre page de [spécifications d’ouverture de l’API](../openspec.md) pour que vous puissiez nous faire part de vos commentaires. Découvrez les nouvelles fonctionnalités du pipeline pour les API JavaScript pour Excel et donnez-nous votre avis sur nos spécifications de conception.

## <a name="excel-javascript-api-reference"></a>Référence de l’API JavaScript pour Excel

Pour plus d’informations sur l’API JavaScript pour Excel, reportez-vous à la rubrique [Documentation de référence des API JavaScript pour Excel](/javascript/api/excel).

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble des compléments Excel](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-overview)
- [Vue d’ensemble de la plateforme des compléments Office](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Exemples de compléments Excel sur GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
