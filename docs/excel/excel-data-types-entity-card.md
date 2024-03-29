---
title: Carte de valeur d’entité des types de données de l’API JavaScript Excel
description: Découvrez comment utiliser des cartes de valeur d’entité avec des types de données dans votre complément Excel.
ms.date: 10/17/2022
ms.topic: conceptual
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 1cb6c49e0e8cb07afb4b7c78a360be6c2391437a
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607568"
---
# <a name="use-cards-with-entity-value-data-types"></a>Utiliser des cartes avec des types de données de valeur d’entité

Cet article explique comment utiliser [l’API JavaScript Excel](../reference/overview/excel-add-ins-reference-overview.md) pour créer des fenêtres modales de carte dans l’interface utilisateur Excel avec des types de données de valeur d’entité. Ces cartes peuvent afficher des informations supplémentaires contenues dans une valeur d’entité, au-delà de ce qui est déjà visible dans une cellule, telles que les images associées, les informations de catégorie de produit et les attributions de données.

Une valeur d’entité, ou [EntityCellValue](/javascript/api/excel/excel.entitycellvalue), est un conteneur pour les types de données et similaire à un objet dans la programmation orientée objet. Cet article explique comment utiliser les propriétés de carte à valeur d’entité, les options de disposition et les fonctionnalités d’attribution de données pour créer des valeurs d’entité qui s’affichent sous forme de cartes.

La capture d’écran suivante montre un exemple de carte de valeur d’entité ouverte, dans ce cas pour le produit **Tofu** à partir d’une liste de produits d’épicerie.

:::image type="content" source="../images/excel-data-types-entity-card-tofu.png" alt-text="Capture d’écran montrant un type de données de valeur d’entité avec la fenêtre de carte affichée.":::

## <a name="card-properties"></a>Propriétés de la carte

La propriété de valeur [`properties`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-properties-member) d’entité vous permet de définir des informations personnalisées sur vos types de données. La `properties` clé accepte les types de données imbriqués. Chaque propriété imbriqué, ou type de données, doit avoir un et `basicValue` un `type` paramètre.

> [!IMPORTANT]
> Les types de données imbriqués `properties` sont utilisés en combinaison avec les valeurs de [disposition](#card-layout) de carte décrites dans la section suivante de l’article. Après avoir défini un type de données imbriqué, `properties`il doit être affecté dans la `layouts` propriété pour s’afficher sur la carte.

L’extrait de code suivant montre le JSON pour une valeur d’entité avec plusieurs types de données imbriqués dans `properties`.

> [!NOTE]
> Pour expérimenter cet extrait de code JSON dans un exemple complet, ouvrez [Script Lab](../overview/explore-with-script-lab.md) dans Excel et sélectionnez Types de [données : Créer des cartes d’entité à partir de données d’une table](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-entity-values.yaml) de notre bibliothèque **Samples**.

```TypeScript
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        "Product ID": {
            type: Excel.CellValueType.string,
            basicValue: productID.toString() || ""
        },
        "Product Name": {
            type: Excel.CellValueType.string,
            basicValue: productName || ""
        },
        "Image": {
            type: Excel.CellValueType.webImage,
            address: product.productImage || ""
        },
        "Quantity Per Unit": {
            type: Excel.CellValueType.string,
            basicValue: product.quantityPerUnit || ""
        },
        "Unit Price": {
            type: Excel.CellValueType.formattedNumber,
            basicValue: product.unitPrice,
            numberFormat: "$* #,##0.00"
        },
        Discontinued: {
            type: Excel.CellValueType.boolean,
            basicValue: product.discontinued || false
        }
    },
    layouts: {
        // Enter layout settings here.
    }
};
```

La capture d’écran suivante montre une carte de valeur d’entité qui utilise l’extrait de code précédent. La capture d’écran montre les informations **d’ID de produit**, **de nom de produit**, **d’image**, **de quantité par unité** et de **prix unitaire** de l’extrait de code précédent.

:::image type="content" source="../images/excel-data-types-entity-card-properties.png" alt-text="Capture d’écran montrant un type de données de valeur d’entité avec la fenêtre de disposition de carte affichée. La carte affiche le nom du produit, l’ID de produit, la quantité par unité et les informations sur le prix unitaire.":::

### <a name="property-metadata"></a>Métadonnées de propriété

Les propriétés d’entité ont un champ facultatif `propertyMetadata` qui utilise l’objet [`CellValuePropertyMetadata`](/javascript/api/excel/excel.cellvaluepropertymetadata) et offre les propriétés `attribution`, `excludeFrom`et `sublabel`. L’extrait de code suivant montre comment ajouter un `sublabel` élément à la `"Unit Price"` propriété à partir de l’extrait de code précédent. Dans ce cas, la sous-étiquette identifie le type de devise.

> [!NOTE]
> Le `propertyMetadata` champ est disponible uniquement sur les types de données imbriqués dans les propriétés d’entité.

```TypeScript
// This code snippet is an excerpt from the `properties` field of the 
// preceding `EntityCellValue` snippet. "Unit Price" is a property of 
// an entity value.
        "Unit Price": {
            type: Excel.CellValueType.formattedNumber,
            basicValue: product.unitPrice,
            numberFormat: "$* #,##0.00",
            propertyMetadata: {
              sublabel: "USD"
            }
        },
```

La capture d’écran suivante montre une carte de valeur d’entité qui utilise l’extrait de code précédent, affichant les métadonnées `sublabel` de propriété **USD** en regard de la propriété **Prix unitaire** .

:::image type="content" source="../images/excel-data-types-entity-card-property-metadata.png" alt-text="Capture d’écran montrant la sous-étiquette USD en regard du prix unitaire.":::

## <a name="card-layout"></a>Disposition de la carte

La propriété de valeur [`layouts`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-layouts-member) d’entité crée un [`card`](/javascript/api/excel/excel.entityviewlayouts) élément pour l’entité, puis spécifie l’apparence de cette carte, comme le titre de la carte, une image pour la carte et le nombre de sections à afficher.

> [!IMPORTANT]
> Les valeurs imbriqués `layouts` sont utilisées en combinaison avec les types de données [des propriétés de](#card-properties) carte décrits dans la section précédente de l’article. Un type de données imbriqué doit être défini `properties` avant de pouvoir être affecté `layouts` pour s’afficher sur la carte.

Dans la `card` propriété, utilisez l’objet [`CardLayoutStandardProperties`](/javascript/api/excel/excel.cardlayoutstandardproperties) pour définir les composants de la carte comme `title`, `subTitle`et `sections`.

L’extrait de code JSON de valeur d’entité suivant montre une `card` disposition avec des objets imbriqués `title` , `mainImage` ainsi que trois `sections` dans la carte. Notez que la `title` propriété `"Product Name"` a un type de données correspondant dans la section précédente de l’article [sur les propriétés de la carte](#card-properties) . La `mainImage` propriété a également un type de données correspondant `"Image"` dans la section précédente. La `sections` propriété prend un tableau imbriqué et utilise l’objet [`CardLayoutSectionStandardProperties`](/javascript/api/excel/excel.cardlayoutsectionstandardproperties) pour définir l’apparence de chaque section.

Dans chaque section de carte, vous pouvez spécifier des éléments comme `layout`, `title`et `properties`. La `layout` clé utilise l’objet [`CardLayoutListSection`](/javascript/api/excel/excel.cardlayoutlistsection) et accepte la valeur `"List"`. La `properties` clé accepte un tableau de chaînes. Notez que les `properties` valeurs, telles que `"Product ID"`, ont des types de données correspondants dans la section précédente de l’article [sur les propriétés de la carte](#card-properties) . Les sections peuvent également être réductibles et peuvent être définies avec des valeurs booléennes comme réduites ou non réduites lorsque la carte d’entité est ouverte dans l’interface utilisateur Excel.

> [!NOTE]
> Pour expérimenter cet extrait de code JSON dans un exemple complet, ouvrez [Script Lab](../overview/explore-with-script-lab.md) dans Excel et sélectionnez Types de [données : Créer des cartes d’entité à partir de données d’une table](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-entity-values.yaml) de notre bibliothèque **Samples**.

```TypeScript
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        // Enter property settings here.
    },
    layouts: {
        card: {
            title: { 
                property: "Product Name" 
            },
            mainImage: { 
                property: "Image" 
            },
            sections: [
                {
                    layout: "List",
                    properties: ["Product ID"]
                },
                {
                    layout: "List",
                    title: "Quantity and price",
                    collapsible: true,
                    collapsed: false, // This section will not be collapsed when the card is opened.
                    properties: ["Quantity Per Unit", "Unit Price"]
                },
                {
                    layout: "List",
                    title: "Additional information",
                    collapsible: true,
                    collapsed: true, // This section will be collapsed when the card is opened.
                    properties: ["Discontinued"]
                }
            ]
        }
    }
};
```

La capture d’écran suivante montre une carte de valeur d’entité qui utilise les extraits de code précédents. La capture d’écran montre l’objet `mainImage` en haut, suivi de l’objet `title` qui utilise le **nom du produit** et est défini sur **Tofu**. La capture d’écran montre `sections`également . La section **Quantité et prix** est réductible et contient **la quantité par unité** et le **prix unitaire**. Le champ **Informations supplémentaires** est réductible et réduit lorsque la carte est ouverte.

:::image type="content" source="../images/excel-data-types-entity-card-sections.png" alt-text="Capture d’écran montrant un type de données de valeur d’entité avec la fenêtre de disposition de carte affichée. La carte affiche le titre et les sections de la carte.":::

## <a name="card-data-attribution"></a>Attribution de données de carte

Les cartes de valeur d’entité peuvent afficher une attribution de données pour accorder un crédit au fournisseur des informations contenues dans la carte d’entité. La propriété de valeur [`provider`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-provider-member) d’entité utilise l’objet [`CellValueProviderAttributes`](/javascript/api/excel/excel.cellvalueproviderattributes) , qui définit le `description`, `logoSourceAddress`et `logoTargetAddress` les valeurs.

La propriété du fournisseur de données affiche une image dans le coin inférieur gauche de la carte d’entité. Il utilise l’URL `logoSourceAddress` source pour spécifier l’image. La `logoTargetAddress` valeur définit la destination de l’URL si l’image du logo est sélectionnée. La `description` valeur s’affiche sous forme d’info-bulle lorsque vous pointez sur le logo. La `description` valeur s’affiche également sous forme de secours en texte brut si l’adresse `logoSourceAddress` n’est pas définie ou si l’adresse source de l’image est interrompue.

L’extrait de code JSON suivant montre une valeur d’entité qui utilise la `provider` propriété pour spécifier une attribution de fournisseur de données pour l’entité.

> [!NOTE]
> Pour expérimenter cet extrait de code JSON dans un exemple complet, ouvrez [Script Lab](../overview/explore-with-script-lab.md) dans Excel et sélectionnez [Types de données : Propriétés d’attribution de valeur d’entité](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/20-data-types/data-types-entity-attribution.yaml) dans notre bibliothèque **Samples**.

```TypeScript
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        // Enter property settings here.
    },
    layouts: {
        // Enter layout settings here.
    },
    provider: {
        description: product.providerName, // Name of the data provider. Displays as a tooltip when hovering over the logo. Also displays as a fallback if the source address for the image is broken.
        logoSourceAddress: product.sourceAddress, // Source URL of the logo to display.
        logoTargetAddress: product.targetAddress // Destination URL that the logo navigates to when selected.
    }
};
```

La capture d’écran suivante montre une carte de valeur d’entité qui utilise l’extrait de code précédent. La capture d’écran montre l’attribution du fournisseur de données dans le coin inférieur gauche. Dans ce cas, le fournisseur de données est Microsoft et le logo Microsoft s’affiche.

:::image type="content" source="../images/excel-data-types-entity-card-attribution.png" alt-text="Capture d’écran montrant un type de données de valeur d’entité avec la fenêtre de disposition de carte affichée. La carte affiche l’attribution du fournisseur de données dans le coin inférieur gauche.":::

## <a name="next-steps"></a>Prochaines étapes

Essayez l’exemple [Créer et explorer des types de données dans Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-data-types-explorer) dans notre référentiel [OfficeDev/Office-Add-in-samples](https://github.com/OfficeDev/Office-Add-in-samples) . Cet exemple vous guide tout au long de la création, puis du chargement indépendant d’un complément qui crée et modifie des types de données dans un classeur.

## <a name="see-also"></a>Voir aussi

- [Présentation des types de données dans les compléments Excel](excel-data-types-overview.md)
- [Concepts de base des types de données Excel](excel-data-types-concepts.md)
- [Créer et explorer des types de données dans Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-data-types-explorer)
- [Référence de l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)