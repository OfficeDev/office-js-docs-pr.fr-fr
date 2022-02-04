---
title: Excel Concepts fondamentaux des types de données de l’API JavaScript
description: Découvrez les concepts de base pour l’utilisation Excel types de données dans votre Office de données.
ms.date: 01/14/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: 'scenarios:getting-started'
ms.localizationpriority: high
---

# <a name="excel-data-types-core-concepts-preview"></a>Excel concepts fondamentaux des types de données (prévisualisation)

> [!NOTE]
> Les API de types de données sont actuellement disponibles uniquement en prévisualisation publique. L’aperçu API peut être modifiés et n’est pas destinés à utiliser dans un environnement de production. Nous vous recommandons de les tester uniquement dans les environnements de test et de développement. N’utilisez pas un aperçu d’API dans un environnement de production ou dans les documents commerciaux importants.
>
> Pour utiliser les API disponibles en préversion :
>
> - Vous devez référencer la **bibliothèque bêta** sur le réseau de distribution de contenu (CDN) ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) . Le [fichier de définition de](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) type pour la compilation et la IntelliSense TypeScript se trouve aux CDN et [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Vous pouvez installer ces types avec `npm install --save-dev @types/office-js-preview` . Pour plus d’informations, voir le @microsoft du package NPM [office-js.](https://www.npmjs.com/package/@microsoft/office-js)
> - Vous devrez peut-être rejoindre [Office programme Insider pour](https://insider.office.com) accéder à des builds Office plus récentes.
>
> Pour tester les types de données dans Office sur Windows, vous devez avoir un numéro de build Excel supérieur ou égal à 16.0.14626.10000. Pour tester les types de données dans Office sur Mac, vous devez avoir un numéro de build Excel supérieur ou égal à 16.55.21102600.

Cet article explique comment utiliser [l’API JavaScript Excel pour](../reference/overview/excel-add-ins-reference-overview.md) utiliser des types de données. Il présente des concepts fondamentaux pour le développement de types de données.

## <a name="core-concepts"></a>Concepts de base

Utilisez la [`Range.valuesAsJson`](/javascript/api/excel/excel.range#excel-excel-range-valuesasjson-member) propriété pour utiliser des valeurs de type de données. Cette propriété est similaire à [Range.values,](/javascript/api/excel/excel.range#excel-excel-range-values-member)mais renvoie uniquement les quatre types de base : `Range.values` chaîne, nombre, booléen ou valeurs d’erreur. `Range.valuesAsJson` peut renvoyer des informations étendues sur les quatre types de base, et cette propriété peut renvoyer des types de données tels que des valeurs numériques formatées, des entités et des images web.

### <a name="json-schema"></a>Schéma JSON

Chaque type de données utilise un schéma de métadonnées JSON conçu pour ce type. Cela définit le [CellValueType](/javascript/api/excel/excel.cellvaluetype) des données et des informations supplémentaires sur la cellule, telles que `basicValue`, `numberFormat`, ou `address`. Chacune `CellValueType` possède des propriétés disponibles en fonction de ce type. Par exemple, le `webImage` type inclut les [propriétés altText](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-alttext-member) [et attribution.](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-attribution-member) Les sections suivantes montrent des exemples de code JSON pour la valeur numérique mise en forme, la valeur d’entité et les types de données d’image web.

Le schéma de métadonnées JSON pour chaque type de données inclut également une ou plusieurs propriétés en lecture seule qui sont utilisées lorsque les calculs rencontrent des scénarios incompatibles, tels qu’une version d’Excel qui ne répond pas à la condition de numéro de build minimale pour la fonctionnalité des types de données. La propriété `basicType` fait partie des métadonnées JSON de chaque type de données, et il s’agit toujours d’une propriété en lecture seule. La `basicType`propriété est utilisée comme secours lorsque le type de données n’est pas pris en charge ou est formatée de manière incorrecte.

## <a name="formatted-number-values"></a>Valeurs numériques formatées

[L’objet FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) permet aux Excel de définir une propriété `numberFormat` pour une valeur. Une fois affecté, ce format de nombre parcourt les calculs avec la valeur et peut être renvoyé par des fonctions.

L’exemple de code JSON suivant montre le schéma complet d’une valeur numérique mise en forme. La valeur numérique mise en forme`myDate` dans l’exemple de code s’affiche comme **16/16/1990** dans l Excel’interface utilisateur. Si les exigences de compatibilité minimales pour la fonctionnalité de types de données ne sont pas remplies, les calculs utilisent le `basicValue` à la place nombre mis en forme.

```json
// This is an example of the complete JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    basicType: Excel.CellValueType.double, // A readonly property. Used as a fallback in incompatible scenarios.
    numberFormat: "m/d/yyyy"
};
```

## <a name="entity-values"></a>Valeurs d’entité

Une valeur d’entité est un conteneur pour les types de données, semblable à un objet dans la programmation orientée objet. Les entités sont également des tableaux en tant que propriétés d’une valeur d’entité. [L’objet EntityCellValue](/javascript/api/excel/excel.entitycellvalue) permet aux compléments de définir des propriétés telles `type` que , et `text` `properties` . La `properties` propriété permet à la valeur d’entité de définir et de contenir des types de données supplémentaires.

Les propriétés `basicType` et `basicValue`définissent la manière dont les calculs lisent ce type de données d’entité si les exigences de compatibilité minimales pour utiliser les types de données ne sont pas remplies. Dans ce scénario, ce type de données d’entité s’affiche en tant que **#VALUE!** erreur dans l’interface Excel IU.

L’exemple de code JSON suivant montre le schéma complet d’une valeur d’entité qui contient du texte, une image, une date et une valeur de texte supplémentaire.

```json
// This is an example of the complete JSON for an entity value.
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
    }, 
    basicType: Excel.CellValueType.error, // A readonly property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A readonly property. Used as a fallback in incompatible scenarios.
};
```

## <a name="web-image-values"></a>Valeurs d’image Web

[L’objet WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) crée la possibilité de [ stocker une image](#entity-values) dans le cadre d’une entité ou en tant que valeur indépendante dans une plage. Cet objet offre de nombreuses propriétés, notamment `address` `altText` , et `relatedImagesAddress` .

Les propriétés `basicType` et `basicValue` définissent la manière dont les calculs lisent le type de données d’image web si les exigences de compatibilité minimales requises pour utiliser la fonctionnalité des types de données ne sont pas remplies. Dans ce scénario, ce type de données d’image web s’affiche en tant que **#VALUE!** erreur dans l’interface Excel IU.

L’exemple de code JSON suivant montre le schéma complet d’une image web.

```json
// This is an example of the complete JSON for a web image.
const myImage = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw", 
    basicType: Excel.CellValueType.error, // A readonly property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A readonly property. Used as a fallback in incompatible scenarios.
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
- [Fonctions personnalisées et types de données](custom-functions-data-types-concepts.md)