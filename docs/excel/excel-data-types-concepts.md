---
title: Excel Concepts fondamentaux des types de données de l’API JavaScript
description: Découvrez les concepts de base pour l’utilisation Excel types de données dans votre Office de données.
ms.date: 10/14/2022
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 65a69838500733f8be08a15a99baa167a946b82a
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607449"
---
# <a name="excel-data-types-core-concepts"></a>Concepts de base des types de données Excel

Cet article explique comment utiliser [l’API JavaScript Excel pour](../reference/overview/excel-add-ins-reference-overview.md) utiliser des types de données. Il présente des concepts fondamentaux pour le développement de types de données.

## <a name="the-valuesasjson-property"></a>la propriété `valuesAsJson`

La `valuesAsJson` propriété (ou le singulier `valueAsJson` pour [NamedItem](/javascript/api/excel/excel.nameditem)) fait partie intégrante de la création de types de données dans Excel. Cette propriété est une extension des propriétés `values`, telles que [Range.values](/javascript/api/excel/excel.range#excel-excel-range-values-member). Les propriétés`values` et `valuesAsJson` sont utilisées pour accéder à la valeur dans une cellule, mais la propriété `values` retourne uniquement l’un des quatre types de base : chaîne, nombre, booléen ou erreur (sous forme de chaîne). En revanche, `valuesAsJson` retourne des informations développées sur les quatre types de base, et cette propriété peut retourner des types de données tels que des valeurs numériques mises en forme, des entités et des images web.

Les objets suivants proposent la propriété `valuesAsJson`.

- [NamedItem](/javascript/api/excel/excel.nameditem) (as `valueAsJson`)
- [NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)
- [Range](/javascript/api/excel/excel.range)
- [RangeView](/javascript/api/excel/excel.rangeview)
- [TableColumn](/javascript/api/excel/excel.tablecolumn)
- [TableRow](/javascript/api/excel/excel.tablerow)

> [!NOTE]
> Certaines valeurs de cellule changent en fonction des paramètres régionaux d’un utilisateur. La propriété `valuesAsJsonLocal` offre une prise en charge de la localisation et est disponible sur tous les mêmes objets que `valuesAsJson`.

## <a name="cell-values"></a>Valeurs de cellule

La `valuesAsJson` propriété renvoie un alias de type [CellValue](/javascript/api/excel/excel.cellvalue), qui est une [union](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types) des types de données suivants.

- [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)
- [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)
- [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)
- [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)
- [EntityCellValue](/javascript/api/excel/excel.entitycellvalue)
- [ErrorCellValue](/javascript/api/excel/excel.errorcellvalue)
- [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)
- [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)
- [ReferenceCellValue](/javascript/api/excel/excel.referencecellvalue)
- [StringCellValue](/javascript/api/excel/excel.stringcellvalue)
- [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)
- [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)

L’alias de type `CellValue` retourne également l’objet [CellValueExtraProperties](/javascript/api/excel/excel.cellvalueextraproperties), qui est une [intersection](https://www.typescriptlang.org/docs/handbook/2/objects.html#intersection-types) avec le reste des types `*CellValue`. Il ne s’agit pas d’un type de données lui-même. Les propriétés de l’objet `CellValueExtraProperties` sont utilisées avec tous les types de données pour spécifier des détails liés au remplacement des valeurs de cellule.

### <a name="json-schema"></a>Schéma JSON

Chaque type de valeur de cellule retourné par `valuesAsJson` utilise un schéma de métadonnées JSON conçu pour ce type. Outre les propriétés supplémentaires propres à chaque type de données, ces schémas de métadonnées JSON ont tous les propriétés `type`, `basicType`et `basicValue` en commun.

Le `type` définit le [CellValueType](/javascript/api/excel/excel.cellvaluetype) des données. Il `basicType` est toujours en lecture seule et est utilisé comme secours lorsque le type de données n’est pas pris en charge ou est mis en forme de manière incorrecte. Le `basicValue` correspond à la valeur retournée par la propriété `values`. Le `basicValue` est utilisé comme solution de repli lorsque les calculs rencontrent des scénarios incompatibles, tels qu’une version antérieure d’Excel qui ne prend pas en charge la fonctionnalité des types de données. Il `basicValue` est en lecture seule pour `ArrayCellValue`les types de données , `EntityCellValue`, `LinkedEntityCellValue`et `WebImageCellValue` les types de données.

Outre les trois champs que tous les types de données partagent, le schéma de métadonnées JSON pour chaque `*CellValue` a des propriétés disponibles en fonction de ce type. Par exemple, le type [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) inclut les propriétés `altText` et `attribution` , tandis que le type [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) offre les champs `properties` et `text`.

Les sections suivantes montrent des exemples de code JSON pour la valeur numérique mise en forme, la valeur d’entité et les types de données d’image web.

## <a name="formatted-number-values"></a>Valeurs numériques formatées

[L’objet FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) permet aux Excel de définir une propriété `numberFormat` pour une valeur. Une fois affecté, ce format de nombre parcourt les calculs avec la valeur et peut être renvoyé par des fonctions.

L’exemple de code JSON suivant montre le schéma complet d’une valeur numérique mise en forme. La valeur numérique mise en forme`myDate` dans l’exemple de code s’affiche comme **16/16/1990** dans l Excel’interface utilisateur. Si les exigences de compatibilité minimales pour la fonctionnalité de types de données ne sont pas remplies, les calculs utilisent le `basicValue` à la place nombre mis en forme.

```TypeScript
// This is an example of the complete JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate: Excel.FormattedNumberCellValue = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    basicType: Excel.RangeValueType.double, // A read-only property. Used as a fallback in incompatible scenarios.
    numberFormat: "m/d/yyyy"
};
```

Commencez à expérimenter les valeurs de nombre mises en forme en ouvrant [Script Lab](../overview/explore-with-script-lab.md) et en vérifiant les types de [données : extraits de code de nombres mis](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-formatted-number.yaml) en forme dans notre bibliothèque **Samples**.

## <a name="entity-values"></a>Valeurs d’entité

Une valeur d’entité est un conteneur pour les types de données, semblable à un objet dans la programmation orientée objet. Les entités sont également des tableaux en tant que propriétés d’une valeur d’entité. [L’objet EntityCellValue](/javascript/api/excel/excel.entitycellvalue) permet aux compléments de définir des propriétés telles `type` que , et `text` `properties` . La `properties` propriété permet à la valeur d’entité de définir et de contenir des types de données supplémentaires.

Les propriétés `basicType` et `basicValue`définissent la manière dont les calculs lisent ce type de données d’entité si les exigences de compatibilité minimales pour utiliser les types de données ne sont pas remplies. Dans ce scénario, ce type de données d’entité s’affiche en tant que **#VALUE!** erreur dans l’interface Excel IU.

L’exemple de code JSON suivant montre le schéma complet d’une valeur d’entité qui contient du texte, une image, une date et une valeur de texte supplémentaire.

```TypeScript
// This is an example of the complete JSON for an entity value.
// The entity contains text and properties which contain an image, a date, and another text value.
const myEntity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: "A llama",
    properties: {
        image: myImage,
        "start date": myDate,
        "quote": {
            type: Excel.CellValueType.string,
            basicValue: "I love llamas."
        }
    }, 
    basicType: Excel.RangeValueType.error, // A read-only property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A read-only property. Used as a fallback in incompatible scenarios.
};
```

Les valeurs d’entité offrent également une propriété `layouts` qui crée une carte pour l’entité. La carte s’affiche sous forme de fenêtre modale dans l’interface utilisateur Excel et peut afficher des informations supplémentaires contenues dans la valeur de l’entité, au-delà de ce qui est visible dans la cellule. Pour plus d’informations, consultez [Utiliser des cartes avec des types de données de valeur d’entité](excel-data-types-entity-card.md).

Pour explorer les types de données d’entité, commencez par [Script Lab dans Excel](../overview/explore-with-script-lab.md) et [ouvrez les types de données : créez des cartes d’entité à partir de données dans un](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-entity-values.yaml) extrait de table dans notre bibliothèque **Samples**. Types [de données : valeurs d’entité avec références](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-references.yaml) et [types de données :](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-entity-attribution.yaml) les extraits de code des propriétés d’attribution de valeur d’entité offrent un examen plus approfondi des fonctionnalités d’entité.

### <a name="linked-entities"></a>Entités liées

Les valeurs d’entité liées, ou objets [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue) sont un type de valeur d’entité. Ces objets intègrent les données fournies par un service externe et peuvent afficher ces données sous la forme d’une [carte d’entité](excel-data-types-entity-card.md), comme des valeurs d’entité régulières. Les [types de données Actions et Géographie](https://support.microsoft.com/office/excel-data-types-stocks-and-geography-61a33056-9935-484f-8ac8-f1a89e210877) disponibles via l’interface utilisateur Excel sont des valeurs d’entité liées.

## <a name="web-image-values"></a>Valeurs d’image Web

The [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) object creates the ability to store an image as part of an [entity](#entity-values) or as an independent value in a range. This object offers many properties, including `address`, `altText`, and `relatedImagesAddress`.

Les propriétés `basicType` et `basicValue` définissent la manière dont les calculs lisent le type de données d’image web si les exigences de compatibilité minimales requises pour utiliser la fonctionnalité des types de données ne sont pas remplies. Dans ce scénario, ce type de données d’image web s’affiche en tant que **#VALUE!** erreur dans l’interface Excel IU.

L’exemple de code JSON suivant montre le schéma complet d’une image web.

```TypeScript
// This is an example of the complete JSON for a web image.
const myImage: Excel.WebImageCellValue = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw", 
    basicType: Excel.RangeValueType.error, // A read-only property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A read-only property. Used as a fallback in incompatible scenarios.
};
```

Essayez les types de données d’image web en ouvrant [Script Lab](../overview/explore-with-script-lab.md) et en sélectionnant les types de données : extraits de code [d’images web](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-web-image.yaml) dans notre bibliothèque **Samples**.

## <a name="improved-error-support"></a>Amélioration de la prise en charge des erreurs

Les API de types de données exposent les erreurs existantes de l'interface utilisateur d'Excel sous forme d'objets. Maintenant que ces erreurs sont accessibles en tant qu’objets, les compléments peuvent définir ou récupérer des propriétés telles `type`, `errorType` et `errorSubType` .

Voici une liste de tous les objets d’erreur avec prise en charge étendue via les types de données.

- [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)
- [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)
- [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)
- [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)
- [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)
- [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)
- [GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)
- [NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)
- [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)
- [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)
- [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)
- [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)
- [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)
- [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)

Chacun des objets d’erreur peut accéder à une enum via la propriété, et cette enum contient des données supplémentaires `errorSubType` sur l’erreur. Par exemple, `BlockedErrorCellValue` l’objet d’erreur peut accéder à l’enum [BlockedErrorCellValueSubType.](/javascript/api/excel/excel.blockederrorcellvaluesubtype) `BlockedErrorCellValueSubType`L’enum fournit des données supplémentaires sur la cause de l’erreur.

Pour en savoir plus sur les objets d’erreur des types de données, consultez les [types de données : définissez l’extrait de code des valeurs d’erreur](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-error-values.yaml) dans notre bibliothèque **d’exemples** [Script Lab](../overview/explore-with-script-lab.md).

## <a name="next-steps"></a>Prochaines étapes

Découvrez comment les types de données d’entité étendent le potentiel des compléments Excel au-delà d’une grille à 2 dimensions avec l’article [Utiliser des cartes avec des types de données de valeur d’entité](excel-data-types-entity-card.md) .

Utilisez l’exemple [Créer et explorer des types de données dans Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-data-types-explorer) dans notre référentiel [OfficeDev/Office-Add-in-samples](https://github.com/OfficeDev/Office-Add-in-samples) pour expérimenter plus en profondeur les types de données en créant et en chargeant de manière indépendante un complément qui crée et modifie des types de données dans un classeur.

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble des types de données dans Excel de données](excel-data-types-overview.md)
- [Utiliser des cartes avec des types de données de valeur d’entité](excel-data-types-entity-card.md)
- [Créer et explorer des types de données dans Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-data-types-explorer)
- [Fonctions personnalisées et types de données](custom-functions-data-types-concepts.md)
- [Référence de l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)