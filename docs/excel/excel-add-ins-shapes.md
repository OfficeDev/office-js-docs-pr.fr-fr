---
title: Utilisation des formes à l’aide de l’API JavaScript pour Excel
description: Découvrez comment Excel définit les formes comme n’importe quel objet qui se trouve sur la couche de dessin d’Excel.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 7b9a4dba02e28187eeb0f932e245489ca61fcbcc
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609740"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a>Utilisation des formes à l’aide de l’API JavaScript pour Excel

Excel définit les formes comme n’importe quel objet qui se trouve sur la couche de dessin d’Excel. Cela signifie que tout élément en dehors d’une cellule est une forme. Cet article explique comment utiliser des formes géométriques, des lignes et des images conjointement avec les API [Shape](/javascript/api/excel/excel.shape) et [ShapeCollection](/javascript/api/excel/excel.shapecollection) . Les [graphiques](/javascript/api/excel/excel.chart) sont abordés dans leur propre article, en [utilisant des graphiques à l’aide de l’API JavaScript pour Excel](excel-add-ins-charts.md).

L’image suivante montre les formes qui forment un thermomètre.
![Image d’un thermomètre effectuée en tant que forme Excel](../images/excel-shapes.png)

## <a name="create-shapes"></a>Créer des formes

Les formes sont créées et stockées dans la collection Shape d’une feuille de calcul ( `Worksheet.shapes` ). `ShapeCollection`dispose `.add*` de plusieurs méthodes à cet effet. Toutes les formes ont des noms et des ID générés pour ceux-ci lorsqu’ils sont ajoutés à la collection. Il s’agit `name` des `id` Propriétés et, respectivement. `name`peut être défini par votre complément pour une extraction facile avec la `ShapeCollection.getItem(name)` méthode.

Les types de formes suivants sont ajoutés à l’aide de la méthode associée :

| Shape | Add, méthode | Signature |
|-------|------------|-----------|
| Forme géométrique | [addGeometricShape](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| Image (JPEG ou PNG) | [addImage](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-) | `addImage(base64ImageString: string): Excel.Shape` |
| Trait | [addLine](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [addSvg](/javascript/api/excel/excel.shapecollection#addsvg-xml-) | `addSvg(xml: string): Excel.Shape` |
| Zone de texte | [addTextBox](/javascript/api/excel/excel.shapecollection#addtextbox-text-) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a>Formes géométriques

Une forme géométrique est créée avec `ShapeCollection.addGeometricShape` . Cette méthode utilise une énumération [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) comme argument.

L’exemple de code suivant crée un rectangle 150x150 nommé **« Square »** qui est positionné 100 pixels à partir des bords supérieur et gauche de la feuille de calcul.

```js
// This sample creates a rectangle positioned 100 pixels from the top and left sides
// of the worksheet and is 150x150 pixels.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="images"></a>Images

Les images JPEG, PNG et SVG peuvent être insérées dans une feuille de calcul en tant que formes. La `ShapeCollection.addImage` méthode prend une chaîne codée en base64 en tant qu’argument. Il s’agit d’une image JPEG ou PNG sous forme de chaîne. `ShapeCollection.addSvg`prend également une chaîne, bien que cet argument soit un XML qui définit le graphique.

L’exemple de code suivant montre un fichier image en cours de chargement par un [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) sous forme de chaîne. La chaîne contient les métadonnées « base64 » supprimées avant la création de la forme.

```js
// This sample creates an image as a Shape object in the worksheet.
var myFile = document.getElementById("selectedFile");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run(function (context) {
        var startIndex = reader.result.toString().indexOf("base64,");
        var myBase64 = reader.result.toString().substr(startIndex + 7);
        var sheet = context.workbook.worksheets.getItem("MyWorksheet");
        var image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        return context.sync();
    }).catch(errorHandlerFunction);
};

// Read in the image file as a data URL.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="lines"></a>Lines

Une ligne est créée avec `ShapeCollection.addLine` . Cette méthode a besoin des marges gauche et supérieure des points de début et de fin du trait. Il prend également une énumération [ConnectorType](/javascript/api/excel/excel.connectortype) pour spécifier la manière dont la ligne passe d’un point de terminaison à un autre. L’exemple de code suivant crée une ligne droite sur la feuille de calcul.

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

Les lignes peuvent être connectées à d’autres objets Shape. Les `connectBeginShape` `connectEndShape` méthodes et joignent le début et la fin d’une ligne aux formes situées aux points de connexion spécifiés. Les emplacements de ces points varient en fonction de la forme, mais le `Shape.connectionSiteCount` peut être utilisé pour s’assurer que votre complément ne se connecte pas à un point qui est hors limites. Une ligne est déconnectée de toutes les formes attachées à l’aide des `disconnectBeginShape` `disconnectEndShape` méthodes et.

L’exemple de code suivant connecte la ligne **« myLine »** à deux formes nommées **« LeftShape »** et **« RightShape »**.

```js
// This sample connects a line between two shapes at connection points '0' and '3'.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.getItem("MyLine").line;
    line.connectBeginShape(shapes.getItem("LeftShape"), 0);
    line.connectEndShape(shapes.getItem("RightShape"), 3);
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-and-resize-shapes"></a>Déplacer et redimensionner des formes

Les formes sont placées en haut de la feuille de calcul. Leur positionnement est défini par la `left` `top` propriété et. Celles-ci agissent comme des marges des arêtes respectives de la feuille de calcul, avec [0,0] correspondant au coin supérieur gauche. Ces éléments peuvent être définis directement ou ajustés à partir de leur position actuelle à l’aide des `incrementLeft` `incrementTop` méthodes et. Le degré de rotation d’une forme par rapport à la position par défaut est également défini de cette manière, la `rotation` propriété étant la valeur absolue et la `incrementRotation` méthode d’ajustement de la rotation existante.

La profondeur d’une forme par rapport à d’autres formes est définie par la `zorderPosition` propriété. Cette valeur est définie à l’aide de la `setZOrder` méthode, qui prend un [ShapeZOrder](/javascript/api/excel/excel.shapezorder). `setZOrder`ajuste l’ordre de la forme actuelle par rapport aux autres formes.

Votre complément offre plusieurs options permettant de modifier la hauteur et la largeur des formes. La définition de `height` la `width` propriété ou modifie la dimension spécifiée sans modifier l’autre dimension. Le `scaleHeight` et `scaleWidth` ajustez les dimensions respectives de la forme par rapport à la taille actuelle ou d’origine (en fonction de la valeur du [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)fourni). Un paramètre [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) facultatif indique l’emplacement de l’échelle de la forme (angle supérieur gauche, milieu ou inférieur droit). Si la `lockAspectRatio` propriété a la **valeur true**, les méthodes d’étendue gèrent les proportions actuelles de la forme en ajustant également l’autre dimension.

> [!NOTE]
> Les modifications apportées aux `height` `width` Propriétés et affectent uniquement cette propriété, quelle que soit la valeur de la `lockAspectRatio` propriété.

L’exemple de code suivant montre une forme mise à l’horizontale à 1,25 fois sa taille d’origine et pivotée de 30 degrés.

```js
// In this sample, the shape "Octagon" is rotated 30 degrees clockwise
// and scaled 25% larger, with the upper-left corner remaining in place.
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("MyWorksheet");
    var shape = sheet.shapes.getItem("Octagon");
    shape.incrementRotation(30);
    shape.lockAspectRatio = true;
    shape.scaleWidth(
        1.25,
        Excel.ShapeScaleType.currentSize,
        Excel.ShapeScaleFrom.scaleFromTopLeft);
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="text-in-shapes"></a>Texte dans des formes

Les formes géométriques peuvent contenir du texte. Les formes ont une `textFrame` propriété de type [TextFrame](/javascript/api/excel/excel.textframe). L' `TextFrame` objet gère les options d’affichage du texte (par exemple, marges et débordement de texte). `TextFrame.textRange`est un objet [TextRange](/javascript/api/excel/excel.textrange) avec les paramètres Text Content et font.

L’exemple de code suivant crée une forme géométrique appelée « Wave » avec le texte « texte de la forme ». Il ajuste également la forme et les couleurs du texte, ainsi que l’alignement horizontal du texte sur le centre.

```js
// This sample creates a light-blue wave shape and adds the purple text "Shape text" to the center.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var wave = shapes.addGeometricShape(Excel.GeometricShapeType.wave);
    wave.left = 100;
    wave.top = 400;
    wave.height = 50;
    wave.width = 150;
    wave.name = "Wave";
    wave.fill.setSolidColor("lightblue");
    wave.textFrame.textRange.text = "Shape text";
    wave.textFrame.textRange.font.color = "purple";
    wave.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;
    return context.sync();
}).catch(errorHandlerFunction);
```

`addTextBox`Méthode de `ShapeCollection` création d’un `GeometricShape` type `Rectangle` avec un arrière-plan blanc et du texte noir. Il s’agit du même que celui créé par le bouton de la **zone de texte** d’Excel sous l’onglet **insertion** . `addTextBox` prend un argument de chaîne pour définir le texte du `TextRange` .

L’exemple de code suivant illustre la création d’une zone de texte avec le texte « Hello ! ».

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 20;
    textbox.width = 45;
    textbox.name = "Textbox";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="shape-groups"></a>Groupes de formes

Les formes peuvent être regroupées. Cela permet à un utilisateur de les traiter comme une seule entité pour le positionnement, le dimensionnement et d’autres tâches connexes. Un [ShapeGroup](/javascript/api/excel/excel.shapegroup) est un type de `Shape` , donc votre complément traite le groupe comme une seule forme.

L’exemple de code suivant montre trois formes regroupées. L’exemple de code suivant montre que le groupe de formes est déplacé vers la droite de 50 pixels.

```js
// This sample takes three previously-created shapes ("Square", "Pentagon", and "Octagon")
// and groups them into a single ShapeGroup.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var square = shapes.getItem("Square");
    var pentagon = shapes.getItem("Pentagon");
    var octagon = shapes.getItem("Octagon");

    var shapeGroup = shapes.addGroup([square, pentagon, octagon]);
    shapeGroup.name = "Group";
    console.log("Shapes grouped");

    return context.sync();
}).catch(errorHandlerFunction);

// This sample moves the previously created shape group to the right by 50 pixels.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var shapeGroup = sheet.shapes.getItem("Group");
    shapeGroup.incrementLeft(50);
    return context.sync();
}).catch(errorHandlerFunction);
```

> [!IMPORTANT]
> Les formes individuelles au sein du groupe sont référencées par le biais `ShapeGroup.shapes` de la propriété, qui est de type [GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection). Elles ne sont plus accessibles via la collection Shape de la feuille de calcul après avoir été groupées. Par exemple, si votre feuille de calcul comporte trois formes et qu’elles ont toutes été regroupées ensemble, la méthode de la feuille de calcul `shapes.getCount` renvoie un nombre égal à 1.

## <a name="export-shapes-as-images"></a>Exporter des formes en tant qu’images

Tout `Shape` objet peut être converti en image. [Shape. getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) renvoie une chaîne codée en base64. Le format de l’image est spécifié comme un enum [PictureFormat](/javascript/api/excel/excel.pictureformat) transmis à `getAsImage` .

```js
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var shape = sheet.shapes.getItem("Image");
    var stringResult = shape.getAsImage(Excel.PictureFormat.png);

    return context.sync().then(function () {
        console.log(stringResult.value);
        // Instead of logging, your add-in may use the base64-encoded string to save the image as a file or insert it in HTML.
    });
}).catch(errorHandlerFunction);
```

## <a name="delete-shapes"></a>Supprimer des formes

Les formes sont supprimées de la feuille de calcul à l’aide de la `Shape` méthode de l’objet `delete` . Aucune autre métadonnée n’est nécessaire.

L’exemple de code suivant supprime toutes les formes de **MyWorksheet**.

```js
// This deletes all the shapes from "MyWorksheet".
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("MyWorksheet");
    var shapes = sheet.shapes;

    // We'll load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    return context.sync().then(function () {
        shapes.items.forEach(function (shape) {
            shape.delete()
        });
        return context.sync();
    }).catch(errorHandlerFunction);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>Voir aussi

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Utiliser des graphiques à l’aide de l’API JavaScript pour Excel](excel-add-ins-charts.md)
