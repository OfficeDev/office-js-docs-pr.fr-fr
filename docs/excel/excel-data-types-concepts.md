---
title: Excel Concepts fondamentaux des types de données de l’API JavaScript
description: Découvrez les concepts de base pour l’utilisation Excel types de données dans votre Office de données.
ms.date: 11/08/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 6155805245b14d3c3365d759bcd647419266f499
ms.sourcegitcommit: 3d37c42f5e465dac52d231d31717bdbb3bfa0e30
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/10/2021
ms.locfileid: "60889978"
---
# <a name="excel-data-types-core-concepts-preview"></a>Excel concepts fondamentaux des types de données (prévisualisation)

> [!NOTE]
> Les API de types de données sont actuellement disponibles uniquement en prévisualisation publique. L’aperçu API peut être modifiés et n’est pas destinés à utiliser dans un environnement de production. N’utilisez pas un aperçu d’API dans un environnement de production ou dans les documents commerciaux importants.

> [!IMPORTANT]
> Certains des concepts de types de données décrits dans cet article, tels que ceux en cours de développement actif, ne sont pas `Range.valuesAsJSON` encore disponibles en prévisualisation publique. Cet article est conçu comme une introduction conceptuelle. Les concepts décrits dans cet article qui ne sont pas encore en prévisualisation publique seront bientôt publiés en prévisualisation.

Cet article explique comment utiliser [l’API JavaScript Excel pour](../reference/overview/excel-add-ins-reference-overview.md) utiliser des types de données. Il présente des concepts fondamentaux pour le développement de types de données.

## <a name="core-concepts"></a>Concepts de base

Utilisez la `Range.valuesAsJSON` propriété pour utiliser des valeurs de type de données. Cette propriété est similaire à [Range.values,](/javascript/api/excel/excel.range#values)mais renvoie uniquement les quatre types de base : `Range.values` chaîne, nombre, booléen ou valeurs d’erreur. `Range.valuesAsJSON` peut renvoyer des informations étendues sur les quatre types de base, et cette propriété peut renvoyer des types de données tels que des valeurs numériques formatées, des entités et des images web.

### <a name="json-schema"></a>Schéma JSON

Les types de données utilisent un schéma JSON cohérent qui définit [le CellValueType](/javascript/api/excel/excel.cellvaluetype) des données et des informations supplémentaires telles `basicValue` que , ou `numberFormat` `address` . Chacune `CellValueType` possède des propriétés disponibles en fonction de ce type. Par exemple, le `webImage` type inclut les [propriétés altText](/javascript/api/excel/excel.webimagecellvalue#altText) [et attribution.](/javascript/api/excel/excel.webimagecellvalue#attribution) Les sections suivantes montrent des exemples de code JSON pour la valeur numérique mise en forme, la valeur d’entité et les types de données d’image web.

## <a name="formatted-number-values"></a>Valeurs numériques formatées

[L’objet FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) permet aux Excel de définir une propriété `numberFormat` pour une valeur. Une fois affecté, ce format de nombre parcourt les calculs avec la valeur et peut être renvoyé par des fonctions.

L’exemple de code JSON suivant montre une valeur numérique mise en forme. La valeur numérique mise en forme`myDate` dans l’exemple de code s’affiche comme **16/16/1990** dans l Excel’interface utilisateur.

```json
// This is an example of the JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    numberFormat: "m/d/yyyy"
};
```

## <a name="entity-values"></a>Valeurs d’entité

Une valeur d’entité est un conteneur pour les types de données, semblable à un objet dans la programmation orientée objet. Les entités sont également des tableaux en tant que propriétés d’une valeur d’entité. [L’objet EntityCellValue](/javascript/api/excel/excel.entitycellvalue) permet aux compléments de définir des propriétés telles `type` que , et `text` `properties` . La `properties` propriété permet à la valeur d’entité de définir et de contenir des types de données supplémentaires.

L’exemple de code JSON suivant montre une valeur d’entité qui contient du texte, une image, une date et une valeur de texte supplémentaire.

```json
// This is an example of the JSON for an entity value.
// The entity contains text and properties which contain an image, a date, and another text value.
const myEntity = {
    type: Excel.CellValueType.entity,
    text: "A llama",
    properties: {
        image: myImage,
        "start date": myDate,
        "quote": {
            type: Excel.CellValueType.string,
            basicValue: "I love llamas."
        }
    }
};
```

## <a name="web-image-values"></a>Valeurs d’image Web

[L’objet WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) crée la possibilité de [ stocker une image](#entity-values) dans le cadre d’une entité ou en tant que valeur indépendante dans une plage. Cet objet offre de nombreuses propriétés, notamment `address` `altText` , et `relatedImagesAddress` .

L’exemple de code JSON suivant montre comment représenter une image web.

```json
// This is an example of the JSON for a web image.
const myImage = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw"
};
```

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

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble des types de données dans Excel de données](excel-data-types-overview.md)
- [Référence de l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Vue d’ensemble des fonctions personnalisées et des types de données](custom-functions-data-types-overview.md)