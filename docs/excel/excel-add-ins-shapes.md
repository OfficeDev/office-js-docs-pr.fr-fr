---
title: Utiliser des formes à l’aide de l’API JavaScript Excel
description: Découvrez comment Excel définit les formes comme n’importe quel objet qui se trouve sur la couche de dessin d’Excel.
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 507ae05b570e7eef4f3bf5560ca47c1bfbd40f9f
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889596"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a>Utiliser des formes à l’aide de l’API JavaScript Excel

Excel définit les formes comme n’importe quel objet qui se trouve sur la couche de dessin d’Excel. Cela signifie que tout ce qui se trouve en dehors d’une cellule est une forme. Cet article explique comment utiliser des formes géométriques, des lignes et des images conjointement avec les API [Shape](/javascript/api/excel/excel.shape) et [ShapeCollection](/javascript/api/excel/excel.shapecollection) . [Les graphiques](/javascript/api/excel/excel.chart) sont abordés dans leur propre article, [Utiliser des graphiques à l’aide de l’API JavaScript Excel](excel-add-ins-charts.md).

L’image suivante montre des formes qui forment un thermomètre.
![Image d’un thermomètre fait en tant que forme Excel.](../images/excel-shapes.png)

## <a name="create-shapes"></a>Créer des formes

Les formes sont créées et stockées dans la collection de formes d’une feuille de calcul (`Worksheet.shapes`). `ShapeCollection` a plusieurs `.add*` méthodes à cet effet. Toutes les formes ont des noms et des ID générés pour elles lorsqu’elles sont ajoutées à la collection. Il s’agit respectivement des propriétés et `id` des `name` propriétés. `name` peut être défini par votre complément pour une récupération facile avec la `ShapeCollection.getItem(name)` méthode.

Les types de formes suivants sont ajoutés à l’aide de la méthode associée.

| Forme | Add, méthode | Signature |
|-------|------------|-----------|
| Forme géométrique | [addGeometricShape](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgeometricshape-member(1)) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| Image (JPEG ou PNG) | [addImage](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addimage-member(1)) | `addImage(base64ImageString: string): Excel.Shape` |
| Trait | [addLine](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addline-member(1)) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [addSvg](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1)) | `addSvg(xml: string): Excel.Shape` |
| Zone de texte | [addTextBox](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addtextbox-member(1)) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a>Formes géométriques

Une forme géométrique est créée avec `ShapeCollection.addGeometricShape`. Cette méthode prend une énumération [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) comme argument.

L’exemple de code suivant crée un rectangle de 150 x 150 pixels nommé **« Square »** positionné à 100 pixels des côtés supérieur et gauche de la feuille de calcul.

```js
// This sample creates a rectangle positioned 100 pixels from the top and left sides
// of the worksheet and is 150x150 pixels.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;

    let rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";

    await context.sync();
});
```

### <a name="images"></a>Images

Les images JPEG, PNG et SVG peuvent être insérées dans une feuille de calcul sous forme de formes. La `ShapeCollection.addImage` méthode prend une chaîne encodée en base64 comme argument. Il s’agit d’une image JPEG ou PNG sous forme de chaîne. `ShapeCollection.addSvg` prend également une chaîne, bien que cet argument soit XML qui définit le graphique.

L’exemple de code suivant montre un fichier image chargé par un [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) sous la forme d’une chaîne. La chaîne a les métadonnées « base64 », supprimées avant la création de la forme.

```js
// This sample creates an image as a Shape object in the worksheet.
let myFile = document.getElementById("selectedFile");
let reader = new FileReader();

reader.onload = (event) => {
    Excel.run(function (context) {
        let startIndex = reader.result.toString().indexOf("base64,");
        let myBase64 = reader.result.toString().substr(startIndex + 7);
        let sheet = context.workbook.worksheets.getItem("MyWorksheet");
        let image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        return context.sync();
    }).catch(errorHandlerFunction);
};

// Read in the image file as a data URL.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="lines"></a>Lines

Une ligne est créée avec `ShapeCollection.addLine`. Cette méthode a besoin des marges gauche et supérieure des points de début et de fin de la ligne. Il faut également une énumération [ConnectorType](/javascript/api/excel/excel.connectortype) pour spécifier comment la ligne contorte entre les points de terminaison. L’exemple de code suivant crée une ligne droite dans la feuille de calcul.

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    await context.sync();
});
```

Les lignes peuvent être connectées à d’autres objets Shape. Les `connectBeginShape` méthodes et `connectEndShape` le début et la fin d’une ligne sont attachés aux formes aux points de connexion spécifiés. Les emplacements de ces points varient selon la forme, mais ils `Shape.connectionSiteCount` peuvent être utilisés pour vous assurer que votre complément ne se connecte pas à un point hors limites. Une ligne est déconnectée de toutes les formes attachées à l’aide des méthodes et `disconnectEndShape` des `disconnectBeginShape` méthodes.

L’exemple de code suivant connecte la ligne **« MyLine »** à deux formes nommées **« LeftShape »** et **« RightShape ».**

```js
// This sample connects a line between two shapes at connection points '0' and '3'.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let line = shapes.getItem("MyLine").line;
    line.connectBeginShape(shapes.getItem("LeftShape"), 0);
    line.connectEndShape(shapes.getItem("RightShape"), 3);
    await context.sync();
});
```

## <a name="move-and-resize-shapes"></a>Déplacer et redimensionner des formes

Les formes sont assises au-dessus de la feuille de calcul. Leur placement est défini par la propriété et `top` la `left` propriété. Ils agissent comme des marges des bords respectifs de la feuille de calcul, [0, 0] étant le coin supérieur gauche. Ceux-ci peuvent être définis directement ou ajustés à partir de leur position actuelle avec les `incrementLeft` méthodes et `incrementTop` . La rotation d’une forme à partir de la position par défaut est également établie de cette manière, la `rotation` propriété étant la quantité absolue et la `incrementRotation` méthode qui ajuste la rotation existante.

La profondeur d’une forme par rapport à d’autres formes est définie par la `zorderPosition` propriété. Ceci est défini à l’aide de la `setZOrder` méthode, qui prend un [ShapeZOrder](/javascript/api/excel/excel.shapezorder). `setZOrder` ajuste l’ordre de la forme actuelle par rapport aux autres formes.

Votre complément dispose de deux options pour modifier la hauteur et la largeur des formes. La définition de la ou `width` de la `height` propriété modifie la dimension spécifiée sans modifier l’autre dimension. Ajustez `scaleHeight` les `scaleWidth` dimensions respectives de la forme par rapport à la taille actuelle ou d’origine (en fonction de la valeur de [ShapeScaleType](/javascript/api/excel/excel.shapescaletype) fourni). Un paramètre [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) facultatif spécifie à partir de l’endroit où la forme est mise à l’échelle (coin supérieur gauche, milieu ou coin inférieur droit). Si la `lockAspectRatio` propriété l’est `true`, les méthodes d’échelle maintiennent le rapport d’aspect actuel de la forme en ajustant également l’autre dimension.

> [!NOTE]
> Les modifications directes apportées aux propriétés et `width` aux `height` propriétés affectent uniquement cette propriété, quelle que soit la valeur de la `lockAspectRatio` propriété.

L’exemple de code suivant montre une forme mise à l’échelle jusqu’à 1,25 fois sa taille d’origine et pivotée de 30 degrés.

```js
// In this sample, the shape "Octagon" is rotated 30 degrees clockwise
// and scaled 25% larger, with the upper-left corner remaining in place.
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("MyWorksheet");

    let shape = sheet.shapes.getItem("Octagon");
    shape.incrementRotation(30);
    shape.lockAspectRatio = true;
    shape.scaleWidth(
        1.25,
        Excel.ShapeScaleType.currentSize,
        Excel.ShapeScaleFrom.scaleFromTopLeft);

    await context.sync();
});
```

## <a name="text-in-shapes"></a>Texte dans les formes

Les formes géométriques peuvent contenir du texte. Les formes ont une `textFrame` propriété de type [TextFrame](/javascript/api/excel/excel.textframe). L’objet `TextFrame` gère les options d’affichage du texte (telles que les marges et le dépassement de texte). `TextFrame.textRange` est un objet [TextRange](/javascript/api/excel/excel.textrange) avec le contenu du texte et les paramètres de police.

L’exemple de code suivant crée une forme géométrique nommée « Wave » avec le texte « Shape Text ». Il ajuste également les couleurs de forme et de texte, ainsi que définit l’alignement horizontal du texte au centre.

```js
// This sample creates a light-blue wave shape and adds the purple text "Shape text" to the center.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let wave = shapes.addGeometricShape(Excel.GeometricShapeType.wave);
    wave.left = 100;
    wave.top = 400;
    wave.height = 50;
    wave.width = 150;

    wave.name = "Wave";
    wave.fill.setSolidColor("lightblue");

    wave.textFrame.textRange.text = "Shape text";
    wave.textFrame.textRange.font.color = "purple";
    wave.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;

    await context.sync();
});
```

La `addTextBox` méthode de création d’un `ShapeCollection` `GeometricShape` type `Rectangle` avec un arrière-plan blanc et du texte noir. Il s’agit du même que celui créé par le bouton **Zone de texte** d’Excel sous l’onglet **Insertion** . `addTextBox` prend un argument de chaîne pour définir le texte du `TextRange`.

L’exemple de code suivant montre la création d’une zone de texte avec le texte « Hello! ».

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 20;
    textbox.width = 45;
    textbox.name = "Textbox";
    await context.sync();
});
```

## <a name="shape-groups"></a>Groupes de formes

Les formes peuvent être regroupées. Cela permet à un utilisateur de les traiter comme une entité unique pour le positionnement, le dimensionnement et d’autres tâches connexes. Un [Groupe de formes](/javascript/api/excel/excel.shapegroup) est un type de `Shape`, de sorte que votre complément traite le groupe comme une forme unique.

L’exemple de code suivant montre trois formes regroupées. L’exemple de code suivant montre que le groupe de formes est déplacé vers la droite de 50 pixels.

```js
// This sample takes three previously-created shapes ("Square", "Pentagon", and "Octagon")
// and groups them into a single ShapeGroup.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let square = shapes.getItem("Square");
    let pentagon = shapes.getItem("Pentagon");
    let octagon = shapes.getItem("Octagon");

    let shapeGroup = shapes.addGroup([square, pentagon, octagon]);
    shapeGroup.name = "Group";
    console.log("Shapes grouped");

    await context.sync();
});

// This sample moves the previously created shape group to the right by 50 pixels.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let shapeGroup = shapes.getItem("Group");
    shapeGroup.incrementLeft(50);
    await context.sync();
});
```

> [!IMPORTANT]
> Les formes individuelles du groupe sont référencées via la `ShapeGroup.shapes` propriété, qui est de type [GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection). Elles ne sont plus accessibles par le biais de la collection de formes de la feuille de calcul après avoir été regroupées. Par exemple, si votre feuille de calcul comporte trois formes et qu’elles sont toutes regroupées, la méthode de `shapes.getCount` la feuille de calcul renvoie un nombre de 1.

## <a name="export-shapes-as-images"></a>Exporter des formes en tant qu’images

N’importe quel `Shape` objet peut être converti en image. [Shape.getAsImage](/javascript/api/excel/excel.shape#excel-excel-shape-getasimage-member(1)) retourne une chaîne encodée en base64. Le format de l’image est spécifié en tant qu’énumération [PictureFormat](/javascript/api/excel/excel.pictureformat) passée à `getAsImage`.

```js
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let shape = shapes.getItem("Image");
    let stringResult = shape.getAsImage(Excel.PictureFormat.png);

    await context.sync();

    console.log(stringResult.value);
    // Instead of logging, your add-in may use the base64-encoded string to save the image as a file or insert it in HTML.
});
```

## <a name="delete-shapes"></a>Supprimer des formes

Les formes sont supprimées de la feuille de calcul avec la méthode de l’objet `Shape` `delete` . Aucune autre métadonnée n’est nécessaire.

L’exemple de code suivant supprime toutes les formes de **MyWorksheet**.

```js
// This deletes all the shapes from "MyWorksheet".
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("MyWorksheet");
    let shapes = sheet.shapes;

    // We'll load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    await context.sync();

    shapes.items.forEach(function (shape) {
        shape.delete();
    });
    
    await context.sync();
});
```

## <a name="see-also"></a>Voir aussi

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Utiliser des graphiques à l’aide de l’API JavaScript pour Excel](excel-add-ins-charts.md)
