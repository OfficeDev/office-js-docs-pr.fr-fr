---
title: Utiliser des formes à l’aide PowerPoint API JavaScript
description: Découvrez comment ajouter, supprimer et mettre en forme des formes sur PowerPoint diapositives.
ms.date: 02/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2c7eb7a1770f807878320369951faa7d0ddc873c
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340483"
---
# <a name="work-with-shapes-using-the-powerpoint-javascript-api-preview"></a>Utiliser des formes à l’aide PowerPoint’API JavaScript (aperçu)

Cet article explique comment utiliser des formes géométriques, des lignes et des zones de texte conjointement avec les API [Shape](/javascript/api/powerpoint/powerpoint.shape) et [ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection) .

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="create-shapes"></a>Créer des formes

Les formes sont créées et stockées dans la collection de formes d’une diapositive (`slide.shapes`). `ShapeCollection` a plusieurs `.add*` méthodes à cet effet. Toutes les formes ont des noms et des ID générés pour elles lorsqu’elles sont ajoutées à la collection. Ce sont respectivement `name` les propriétés `id` et les propriétés. `name` peut être définie par votre add-in.

### <a name="geometric-shapes"></a>Formes géométriques

Une forme géométrique est créée avec l’une des surcharges de `ShapeCollection.addGeometricShape`. Le premier paramètre est une [enum GeometricShapeType](/javascript/api/powerpoint/powerpoint.geometricshapetype) ou l’équivalent de chaîne de l’une des valeurs de l’enum. Il existe un deuxième paramètre facultatif de type [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) qui peut spécifier la taille initiale de la forme et sa position par rapport aux côtés supérieur et gauche de la diapositive, mesurée en points. Ou ces propriétés peuvent être définies après la création de la forme.

L’exemple de code suivant crée un rectangle nommé « **Square** » placé à 100 points des côtés supérieur et gauche de la diapositive. La méthode renvoie un `Shape` objet.

```js
// This sample creates a rectangle positioned 100 points from the top and left sides
// of the slide and is 150x150 points. The shape is put on the first slide.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const rectangle = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";
    await context.sync();
});
```

### <a name="lines"></a>Lines

Une ligne est créée avec l’une des surcharges de `ShapeCollection.addLine`. Le premier paramètre est une enum [ConnectorType](/javascript/api/powerpoint/powerpoint.connectortype) ou l’équivalent de chaîne d’une des valeurs de l’enum pour spécifier la façon dont la ligne se contorte entre les points de terminaison. Il existe un deuxième paramètre facultatif de type [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) qui peut spécifier les points de début et de fin de la ligne. Ou ces propriétés peuvent être définies après la création de la forme. La méthode renvoie un `Shape` objet.

> [!NOTE]
> Lorsque la forme est une ligne, `top` `left` `Shape` `ShapeAddOptions` les propriétés et les objets spécifient le point de départ de la ligne par rapport aux bords supérieur et gauche de la diapositive. Les `height` propriétés `width` spécifient le point de terminaison de la ligne *par rapport au point de départ*. Ainsi, le point de fin par rapport aux bords supérieur et gauche de la diapositive est (`top` + `height`) de ().`left` + `width` L’unité de mesure de toutes les propriétés est de points et les valeurs négatives sont autorisées.

L’exemple de code suivant crée une ligne droite sur la diapositive.

```js
// This sample creates a straight line on the first slide.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const line = shapes.addLine(Excel.ConnectorType.straight, {left: 200, top: 50, height: 300, width: 150});
    line.name = "StraightLine";
    await context.sync();
});
```

### <a name="text-boxes"></a>Zones de texte

Une zone de texte est créée avec la [méthode addTextBox](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addtextbox-member(1)) . Le premier paramètre est le texte qui doit apparaître dans la zone initialement. Il existe un deuxième paramètre facultatif de type [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) qui peut spécifier la taille initiale de la zone de texte et sa position par rapport aux côtés supérieur et gauche de la diapositive. Ou ces propriétés peuvent être définies après la création de la forme.

L’exemple de code suivant montre comment créer une zone de texte sur la première diapositive.

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 300;
    textbox.width = 450;
    textbox.name = "Textbox";
    await context.sync();
});
```

## <a name="move-and-resize-shapes"></a>Déplacer et re tailler des formes

Les formes sont au-dessus de la diapositive. Leur placement est défini par les propriétés `left` et les `top` propriétés. Elles agissent comme des marges des bords respectifs de la diapositive, mesurées en points, `left: 0` `top: 0` avec et en étant le coin supérieur gauche. La taille de la forme est spécifiée par les propriétés `height` et les `width` propriétés. Votre code peut déplacer ou reizer la forme en réinitialisation de ces propriétés. (Ces propriétés ont une signification légèrement différente lorsque la forme est un trait. Voir [Lignes](#lines).)

## <a name="text-in-shapes"></a>Texte dans les formes

Les formes géométriques peuvent contenir du texte. Les formes ont `textFrame` une propriété de type [TextFrame](/javascript/api/powerpoint/powerpoint.textframe). L’objet `TextFrame` gère les options d’affichage de texte (telles que les marges et le dépassement de texte). `TextFrame.textRange` est un [objet TextRange](/javascript/api/powerpoint/powerpoint.textrange) avec le contenu du texte et les paramètres de police.

L’exemple de code suivant crée une forme géométrique nommée **« Braces** » avec le texte **« Shape text** ». Il ajuste également les couleurs de la forme et du texte, et définit l’alignement vertical du texte sur le centre.

```js
// This sample creates a light blue rectangle with braces ("{}") on the left and right ends
// and adds the purple text "Shape text" to the center.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const braces = shapes.addGeometricShape(PowerPoint.GeometricShapeType.bracePair);
    braces.left = 100;
    braces.top = 400;
    braces.height = 50;
    braces.width = 150;
    braces.name = "Braces";
    braces.fill.setSolidColor("lightblue");
    braces.textFrame.textRange.text = "Shape text";
    braces.textFrame.textRange.font.color = "purple";
    braces.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
    await context.sync();
});
```

## <a name="delete-shapes"></a>Supprimer des formes

Les formes sont supprimées de la diapositive à l’aide `Shape` de la méthode de l’objet `delete` .

L’exemple de code suivant montre comment supprimer des formes.

```js
await PowerPoint.run(async (context) => {
    // Delete all shapes from the first slide.
    const sheet = context.presentation.slides.getItemAt(0);
    const shapes = sheet.shapes;

    // Load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    await context.sync();
        
    shapes.items.forEach(function (shape) {
        shape.delete();
    });
    await context.sync();
});
```
