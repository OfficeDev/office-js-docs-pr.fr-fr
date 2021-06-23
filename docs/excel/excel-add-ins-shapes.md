---
title: Utiliser des formes à l’aide Excel API JavaScript
description: Découvrez comment Excel définit les formes comme n’importe quel objet qui se trouve sur la couche de dessin de Excel.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 936def11a5d597b68cc59a58b041c4f30ff46a38
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075760"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a><span data-ttu-id="5727f-103">Utiliser des formes à l’aide Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="5727f-103">Work with shapes using the Excel JavaScript API</span></span>

<span data-ttu-id="5727f-104">Excel définit les formes comme n’importe quel objet qui se trouve sur la couche de dessin de Excel.</span><span class="sxs-lookup"><span data-stu-id="5727f-104">Excel defines shapes as any object that sits on the drawing layer of Excel.</span></span> <span data-ttu-id="5727f-105">Cela signifie que tout ce qui se trouve en dehors d’une cellule est une forme.</span><span class="sxs-lookup"><span data-stu-id="5727f-105">That means anything outside of a cell is a shape.</span></span> <span data-ttu-id="5727f-106">Cet article explique comment utiliser des formes géométriques, des lignes et des images conjointement avec les API [Shape](/javascript/api/excel/excel.shape) et [ShapeCollection.](/javascript/api/excel/excel.shapecollection)</span><span class="sxs-lookup"><span data-stu-id="5727f-106">This article describes how to use geometric shapes, lines, and images in conjunction with the [Shape](/javascript/api/excel/excel.shape) and [ShapeCollection](/javascript/api/excel/excel.shapecollection) APIs.</span></span> <span data-ttu-id="5727f-107">[Les](/javascript/api/excel/excel.chart) graphiques sont abordés dans leur propre article, [Work with charts using the Excel JavaScript API](excel-add-ins-charts.md).</span><span class="sxs-lookup"><span data-stu-id="5727f-107">[Charts](/javascript/api/excel/excel.chart) are covered in their own article, [Work with charts using the Excel JavaScript API](excel-add-ins-charts.md).</span></span>

<span data-ttu-id="5727f-108">L’image suivante montre les formes qui forment un thermomètre.</span><span class="sxs-lookup"><span data-stu-id="5727f-108">The following image shows shapes which form a thermometer.</span></span>
<span data-ttu-id="5727f-109">![Image d’un thermomètre en forme de Excel forme.](../images/excel-shapes.png)</span><span class="sxs-lookup"><span data-stu-id="5727f-109">![Image of a thermometer made as an Excel shape.](../images/excel-shapes.png)</span></span>

## <a name="create-shapes"></a><span data-ttu-id="5727f-110">Créer des formes</span><span class="sxs-lookup"><span data-stu-id="5727f-110">Create shapes</span></span>

<span data-ttu-id="5727f-111">Les formes sont créées par le biais et stockées dans la collection de formes d’une feuille de calcul ( `Worksheet.shapes` ).</span><span class="sxs-lookup"><span data-stu-id="5727f-111">Shapes are created through and stored in a worksheet's shape collection (`Worksheet.shapes`).</span></span> <span data-ttu-id="5727f-112">`ShapeCollection` a plusieurs `.add*` méthodes à cet effet.</span><span class="sxs-lookup"><span data-stu-id="5727f-112">`ShapeCollection` has several `.add*` methods for this purpose.</span></span> <span data-ttu-id="5727f-113">Toutes les formes ont des noms et des ID générés pour elles lorsqu’elles sont ajoutées à la collection.</span><span class="sxs-lookup"><span data-stu-id="5727f-113">All shapes have names and IDs generated for them when they are added to the collection.</span></span> <span data-ttu-id="5727f-114">Ce sont `name` respectivement les `id` propriétés et les propriétés.</span><span class="sxs-lookup"><span data-stu-id="5727f-114">These are the `name` and `id` properties, respectively.</span></span> <span data-ttu-id="5727f-115">`name` peut être définie par votre add-in pour faciliter l’extraction avec la `ShapeCollection.getItem(name)` méthode.</span><span class="sxs-lookup"><span data-stu-id="5727f-115">`name` can be set by your add-in for easy retrieval with the `ShapeCollection.getItem(name)` method.</span></span>

<span data-ttu-id="5727f-116">Les types de formes suivants sont ajoutés à l’aide de la méthode associée :</span><span class="sxs-lookup"><span data-stu-id="5727f-116">The following types of shapes are added using the associated method:</span></span>

| <span data-ttu-id="5727f-117">Forme</span><span class="sxs-lookup"><span data-stu-id="5727f-117">Shape</span></span> | <span data-ttu-id="5727f-118">Add, méthode</span><span class="sxs-lookup"><span data-stu-id="5727f-118">Add Method</span></span> | <span data-ttu-id="5727f-119">Signature</span><span class="sxs-lookup"><span data-stu-id="5727f-119">Signature</span></span> |
|-------|------------|-----------|
| <span data-ttu-id="5727f-120">Forme géométrique</span><span class="sxs-lookup"><span data-stu-id="5727f-120">Geometric Shape</span></span> | [<span data-ttu-id="5727f-121">addGeometricShape</span><span class="sxs-lookup"><span data-stu-id="5727f-121">addGeometricShape</span></span>](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| <span data-ttu-id="5727f-122">Image (JPEG ou PNG)</span><span class="sxs-lookup"><span data-stu-id="5727f-122">Image (either JPEG or PNG)</span></span> | [<span data-ttu-id="5727f-123">addImage</span><span class="sxs-lookup"><span data-stu-id="5727f-123">addImage</span></span>](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-) | `addImage(base64ImageString: string): Excel.Shape` |
| <span data-ttu-id="5727f-124">Trait</span><span class="sxs-lookup"><span data-stu-id="5727f-124">Line</span></span> | [<span data-ttu-id="5727f-125">addLine</span><span class="sxs-lookup"><span data-stu-id="5727f-125">addLine</span></span>](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| <span data-ttu-id="5727f-126">SVG</span><span class="sxs-lookup"><span data-stu-id="5727f-126">SVG</span></span> | [<span data-ttu-id="5727f-127">addSvg</span><span class="sxs-lookup"><span data-stu-id="5727f-127">addSvg</span></span>](/javascript/api/excel/excel.shapecollection#addsvg-xml-) | `addSvg(xml: string): Excel.Shape` |
| <span data-ttu-id="5727f-128">Zone de texte</span><span class="sxs-lookup"><span data-stu-id="5727f-128">Text Box</span></span> | [<span data-ttu-id="5727f-129">addTextBox</span><span class="sxs-lookup"><span data-stu-id="5727f-129">addTextBox</span></span>](/javascript/api/excel/excel.shapecollection#addtextbox-text-) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a><span data-ttu-id="5727f-130">Formes géométriques</span><span class="sxs-lookup"><span data-stu-id="5727f-130">Geometric shapes</span></span>

<span data-ttu-id="5727f-131">Une forme géométrique est créée avec `ShapeCollection.addGeometricShape` .</span><span class="sxs-lookup"><span data-stu-id="5727f-131">A geometric shape is created with `ShapeCollection.addGeometricShape`.</span></span> <span data-ttu-id="5727f-132">Cette méthode prend une [enum GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) comme argument.</span><span class="sxs-lookup"><span data-stu-id="5727f-132">That method takes a [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) enum as an argument.</span></span>

<span data-ttu-id="5727f-133">L’exemple de code suivant crée un rectangle de 150 x 150 pixels nommé « **Square** » placé à 100 pixels des côtés supérieur et gauche de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="5727f-133">The following code sample creates a 150x150-pixel rectangle named **"Square"** that is positioned 100 pixels from the top and left sides of the worksheet.</span></span>

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

### <a name="images"></a><span data-ttu-id="5727f-134">Des images</span><span class="sxs-lookup"><span data-stu-id="5727f-134">Images</span></span>

<span data-ttu-id="5727f-135">Les images JPEG, PNG et SVG peuvent être insérées dans une feuille de calcul sous forme de formes.</span><span class="sxs-lookup"><span data-stu-id="5727f-135">JPEG, PNG, and SVG images can be inserted into a worksheet as shapes.</span></span> <span data-ttu-id="5727f-136">La méthode prend comme argument une chaîne `ShapeCollection.addImage` codée en base 64.</span><span class="sxs-lookup"><span data-stu-id="5727f-136">The `ShapeCollection.addImage` method takes a base64-encoded string as an argument.</span></span> <span data-ttu-id="5727f-137">Il s’agit d’une image JPEG ou PNG sous forme de chaîne.</span><span class="sxs-lookup"><span data-stu-id="5727f-137">This is either a JPEG or PNG image in string form.</span></span> <span data-ttu-id="5727f-138">`ShapeCollection.addSvg` prend également une chaîne, bien que cet argument soit XML qui définit le graphique.</span><span class="sxs-lookup"><span data-stu-id="5727f-138">`ShapeCollection.addSvg` also takes in a string, though this argument is XML that defines the graphic.</span></span>

<span data-ttu-id="5727f-139">L’exemple de code suivant montre un fichier image chargé par [un FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) sous la mesure d’une chaîne.</span><span class="sxs-lookup"><span data-stu-id="5727f-139">The following code sample shows an image file being loaded by a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) as a string.</span></span> <span data-ttu-id="5727f-140">La chaîne a les métadonnées « base64 » supprimées avant la création de la forme.</span><span class="sxs-lookup"><span data-stu-id="5727f-140">The string has the metadata "base64," removed before the shape is created.</span></span>

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

### <a name="lines"></a><span data-ttu-id="5727f-141">Lines</span><span class="sxs-lookup"><span data-stu-id="5727f-141">Lines</span></span>

<span data-ttu-id="5727f-142">Une ligne est créée avec `ShapeCollection.addLine` .</span><span class="sxs-lookup"><span data-stu-id="5727f-142">A line is created with `ShapeCollection.addLine`.</span></span> <span data-ttu-id="5727f-143">Cette méthode a besoin des marges gauche et supérieure des points de début et de fin de la ligne.</span><span class="sxs-lookup"><span data-stu-id="5727f-143">That method needs the left and top margins of the line's start and end points.</span></span> <span data-ttu-id="5727f-144">Il faut également une enum [ConnectorType](/javascript/api/excel/excel.connectortype) pour spécifier la façon dont la ligne se contorte entre les points de terminaison.</span><span class="sxs-lookup"><span data-stu-id="5727f-144">It also takes a [ConnectorType](/javascript/api/excel/excel.connectortype) enum to specify how the line contorts between endpoints.</span></span> <span data-ttu-id="5727f-145">L’exemple de code suivant crée une ligne droite sur la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="5727f-145">The following code sample creates a straight line on the worksheet.</span></span>

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="5727f-146">Les lignes peuvent être connectées à d’autres objets Shape.</span><span class="sxs-lookup"><span data-stu-id="5727f-146">Lines can be connected to other Shape objects.</span></span> <span data-ttu-id="5727f-147">Les méthodes attachent le début et la fin d’une ligne aux formes aux `connectBeginShape` points de connexion `connectEndShape` spécifiés.</span><span class="sxs-lookup"><span data-stu-id="5727f-147">The `connectBeginShape` and `connectEndShape` methods attach the beginning and ending of a line to shapes at the specified connection points.</span></span> <span data-ttu-id="5727f-148">Les emplacements de ces points varient en fonction de la forme, mais ils peuvent être utilisés pour vous assurer que votre module ne se connecte pas à un point hors `Shape.connectionSiteCount` limites.</span><span class="sxs-lookup"><span data-stu-id="5727f-148">The locations of these points vary by shape, but the `Shape.connectionSiteCount` can be used to ensure your add-in does not connect to a point that's out-of-bounds.</span></span> <span data-ttu-id="5727f-149">Une ligne est déconnectée des formes attachées à l’aide `disconnectBeginShape` des méthodes et des `disconnectEndShape` formes.</span><span class="sxs-lookup"><span data-stu-id="5727f-149">A line is disconnected from any attached shapes using the `disconnectBeginShape` and `disconnectEndShape` methods.</span></span>

<span data-ttu-id="5727f-150">L’exemple de code suivant connecte la ligne « **MyLine** » à deux formes nommées **« LeftShape** » et **« RightShape**».</span><span class="sxs-lookup"><span data-stu-id="5727f-150">The following code sample connects the **"MyLine"** line to two shapes named **"LeftShape"** and **"RightShape"**.</span></span>

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

## <a name="move-and-resize-shapes"></a><span data-ttu-id="5727f-151">Déplacer et re tailler des formes</span><span class="sxs-lookup"><span data-stu-id="5727f-151">Move and resize shapes</span></span>

<span data-ttu-id="5727f-152">Les formes sont au-dessus de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="5727f-152">Shapes sit on top of the worksheet.</span></span> <span data-ttu-id="5727f-153">Leur placement est défini par la `left` propriété `top` et la propriété.</span><span class="sxs-lookup"><span data-stu-id="5727f-153">Their placement is defined by the `left` and `top` property.</span></span> <span data-ttu-id="5727f-154">Elles agissent comme des marges des bords respectifs de la feuille de calcul, avec [0, 0] en tant que coin supérieur gauche.</span><span class="sxs-lookup"><span data-stu-id="5727f-154">These act as margins from worksheet's respective edges, with [0, 0] being the upper-left corner.</span></span> <span data-ttu-id="5727f-155">Celles-ci peuvent être définies directement ou ajustées à partir de leur position actuelle avec les `incrementLeft` méthodes `incrementTop` et les méthodes.</span><span class="sxs-lookup"><span data-stu-id="5727f-155">These can either be set directly or adjusted from their current position with the `incrementLeft` and `incrementTop` methods.</span></span> <span data-ttu-id="5727f-156">La quantité de rotation d’une forme par rapport à la position par défaut est également établie de cette manière, la propriété étant la quantité absolue et la méthode ajustant la `rotation` `incrementRotation` rotation existante.</span><span class="sxs-lookup"><span data-stu-id="5727f-156">How much a shape is rotated from the default position is also established in this manner, with the `rotation` property being the absolute amount and the `incrementRotation` method adjusting the existing rotation.</span></span>

<span data-ttu-id="5727f-157">La profondeur d’une forme par rapport aux autres formes est définie par la `zorderPosition` propriété.</span><span class="sxs-lookup"><span data-stu-id="5727f-157">A shape's depth relative to other shapes is defined by the `zorderPosition` property.</span></span> <span data-ttu-id="5727f-158">Ceci est définie à `setZOrder` l’aide de la méthode, qui prend [un ShapeZOrder](/javascript/api/excel/excel.shapezorder).</span><span class="sxs-lookup"><span data-stu-id="5727f-158">This is set using the `setZOrder` method, which takes a [ShapeZOrder](/javascript/api/excel/excel.shapezorder).</span></span> <span data-ttu-id="5727f-159">`setZOrder` ajuste l’ordre de la forme actuelle par rapport aux autres formes.</span><span class="sxs-lookup"><span data-stu-id="5727f-159">`setZOrder` adjusts the ordering of the current shape relative to the other shapes.</span></span>

<span data-ttu-id="5727f-160">Votre add-in dispose de deux options pour modifier la hauteur et la largeur des formes.</span><span class="sxs-lookup"><span data-stu-id="5727f-160">Your add-in has a couple options for changing the height and width of shapes.</span></span> <span data-ttu-id="5727f-161">La définition de `height` la ou de la propriété modifie la dimension `width` spécifiée sans modifier l’autre dimension.</span><span class="sxs-lookup"><span data-stu-id="5727f-161">Setting either the `height` or `width` property changes the specified dimension without changing the other dimension.</span></span> <span data-ttu-id="5727f-162">L’et ajuster les dimensions respectives de la forme par rapport à la taille actuelle ou d’origine (en fonction de la valeur de `scaleHeight` `scaleWidth` [l’shapeScaleType fourni](/javascript/api/excel/excel.shapescaletype)).</span><span class="sxs-lookup"><span data-stu-id="5727f-162">The `scaleHeight` and `scaleWidth` adjust the shape's respective dimensions relative to either the current or original size (based on the value of the provided [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)).</span></span> <span data-ttu-id="5727f-163">Un paramètre [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) facultatif spécifie l’endroit où la forme est mise à l’échelle (coin supérieur gauche, milieu ou coin inférieur droit).</span><span class="sxs-lookup"><span data-stu-id="5727f-163">An optional [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) parameter specifies from where the shape scales (top-left corner, middle, or bottom-right corner).</span></span> <span data-ttu-id="5727f-164">Si la propriété est true, les méthodes d’échelle conservent les proportions actuelles de la forme en ajustant également `lockAspectRatio` l’autre dimension.</span><span class="sxs-lookup"><span data-stu-id="5727f-164">If the `lockAspectRatio` property is **true**, the scale methods maintain the shape's current aspect ratio by also adjusting the other dimension.</span></span>

> [!NOTE]
> <span data-ttu-id="5727f-165">Les modifications directes apportées aux propriétés affectent uniquement cette propriété, quelle que `height` soit la valeur de la `width` `lockAspectRatio` propriété.</span><span class="sxs-lookup"><span data-stu-id="5727f-165">Direct changes to the `height` and `width` properties only affect that property, regardless of the `lockAspectRatio` property's value.</span></span>

<span data-ttu-id="5727f-166">L’exemple de code suivant montre une forme mise à l’échelle jusqu’à 1,25 fois sa taille d’origine et pivotée de 30 degrés.</span><span class="sxs-lookup"><span data-stu-id="5727f-166">The following code sample shows a shape being scaled to 1.25 times its original size and rotated 30 degrees.</span></span>

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

## <a name="text-in-shapes"></a><span data-ttu-id="5727f-167">Texte dans les formes</span><span class="sxs-lookup"><span data-stu-id="5727f-167">Text in shapes</span></span>

<span data-ttu-id="5727f-168">Les formes géométriques peuvent contenir du texte.</span><span class="sxs-lookup"><span data-stu-id="5727f-168">Geometric Shapes can contain text.</span></span> <span data-ttu-id="5727f-169">Les formes ont `textFrame` une propriété de type [TextFrame](/javascript/api/excel/excel.textframe).</span><span class="sxs-lookup"><span data-stu-id="5727f-169">Shapes have a `textFrame` property of type [TextFrame](/javascript/api/excel/excel.textframe).</span></span> <span data-ttu-id="5727f-170">`TextFrame`L’objet gère les options d’affichage de texte (telles que les marges et le dépassement de texte).</span><span class="sxs-lookup"><span data-stu-id="5727f-170">The `TextFrame` object manages the text display options (such as margins and text overflow).</span></span> <span data-ttu-id="5727f-171">`TextFrame.textRange` est un [objet TextRange](/javascript/api/excel/excel.textrange) avec le contenu du texte et les paramètres de police.</span><span class="sxs-lookup"><span data-stu-id="5727f-171">`TextFrame.textRange` is a [TextRange](/javascript/api/excel/excel.textrange) object with the text content and font settings.</span></span>

<span data-ttu-id="5727f-172">L’exemple de code suivant crée une forme géométrique nommée « Wave » avec le texte « Shape Text ».</span><span class="sxs-lookup"><span data-stu-id="5727f-172">The following code sample creates a geometric shape named "Wave" with the text "Shape Text".</span></span> <span data-ttu-id="5727f-173">Il ajuste également les couleurs de la forme et du texte, et définit l’alignement horizontal du texte au centre.</span><span class="sxs-lookup"><span data-stu-id="5727f-173">It also adjusts the shape and text colors, as well as sets the text's horizontal alignment to the center.</span></span>

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

<span data-ttu-id="5727f-174">La `addTextBox` méthode de création `ShapeCollection` `GeometricShape` d’un type avec un `Rectangle` arrière-plan blanc et du texte noir.</span><span class="sxs-lookup"><span data-stu-id="5727f-174">The `addTextBox` method of `ShapeCollection` creates a `GeometricShape` of type `Rectangle` with a white background and black text.</span></span> <span data-ttu-id="5727f-175">Ceci est identique à ce qui est créé par Excel bouton **Zone** de texte sous **l’onglet** Insertion. `addTextBox` prend un argument de chaîne pour définir le texte du `TextRange` .</span><span class="sxs-lookup"><span data-stu-id="5727f-175">This is the same as what is created by Excel's **Text Box** button on the **Insert** tab. `addTextBox` takes a string argument to set the text of the `TextRange`.</span></span>

<span data-ttu-id="5727f-176">L’exemple de code suivant montre la création d’une zone de texte avec le texte « Hello! ».</span><span class="sxs-lookup"><span data-stu-id="5727f-176">The following code sample shows the creation of a text box with the text "Hello!".</span></span>

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

## <a name="shape-groups"></a><span data-ttu-id="5727f-177">Groupes de formes</span><span class="sxs-lookup"><span data-stu-id="5727f-177">Shape groups</span></span>

<span data-ttu-id="5727f-178">Les formes peuvent être regroupées.</span><span class="sxs-lookup"><span data-stu-id="5727f-178">Shapes can be grouped together.</span></span> <span data-ttu-id="5727f-179">Cela permet à un utilisateur de les traiter comme une entité unique pour le positionnement, le resserrement et d’autres tâches connexes.</span><span class="sxs-lookup"><span data-stu-id="5727f-179">This allows a user to treat them as a single entity for positioning, sizing, and other related tasks.</span></span> <span data-ttu-id="5727f-180">Un [shapeGroup](/javascript/api/excel/excel.shapegroup) est un type de `Shape` , donc votre add-in traite le groupe comme une forme unique.</span><span class="sxs-lookup"><span data-stu-id="5727f-180">A [ShapeGroup](/javascript/api/excel/excel.shapegroup) is a type of `Shape`, so your add-in treats the group as a single shape.</span></span>

<span data-ttu-id="5727f-181">L’exemple de code suivant montre trois formes regroupées.</span><span class="sxs-lookup"><span data-stu-id="5727f-181">The following code sample shows three shapes being grouped together.</span></span> <span data-ttu-id="5727f-182">L’exemple de code suivant montre que le groupe de formes est déplacé vers la droite de 50 pixels.</span><span class="sxs-lookup"><span data-stu-id="5727f-182">The subsequent code sample shows that shape group being moved to the right 50 pixels.</span></span>

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
> <span data-ttu-id="5727f-183">Les formes individuelles au sein du groupe sont référencés par le biais de la `ShapeGroup.shapes` propriété, qui est de type [GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection).</span><span class="sxs-lookup"><span data-stu-id="5727f-183">Individual shapes within the group are referenced through the `ShapeGroup.shapes` property, which is of type [GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection).</span></span> <span data-ttu-id="5727f-184">Elles ne sont plus accessibles par le biais de la collection de formes de la feuille de calcul après avoir été regroupées.</span><span class="sxs-lookup"><span data-stu-id="5727f-184">They are no longer accessible through the worksheet's shape collection after being grouped.</span></span> <span data-ttu-id="5727f-185">Par exemple, si votre feuille de calcul avait trois formes et qu’elles étaient toutes regroupées, la méthode de la feuille de calcul retournerait le nombre `shapes.getCount` 1.</span><span class="sxs-lookup"><span data-stu-id="5727f-185">As an example, if your worksheet had three shapes and they were all grouped together, the worksheet's `shapes.getCount` method would return a count of 1.</span></span>

## <a name="export-shapes-as-images"></a><span data-ttu-id="5727f-186">Exporter des formes en tant qu’images</span><span class="sxs-lookup"><span data-stu-id="5727f-186">Export shapes as images</span></span>

<span data-ttu-id="5727f-187">Tout `Shape` objet peut être converti en image.</span><span class="sxs-lookup"><span data-stu-id="5727f-187">Any `Shape` object can be converted to an image.</span></span> <span data-ttu-id="5727f-188">[Shape.getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) renvoie une chaîne codée en base 64.</span><span class="sxs-lookup"><span data-stu-id="5727f-188">[Shape.getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) returns base64-encoded string.</span></span> <span data-ttu-id="5727f-189">Le format de l’image est spécifié en tant qu’enum [PictureFormat](/javascript/api/excel/excel.pictureformat) transmis à `getAsImage` .</span><span class="sxs-lookup"><span data-stu-id="5727f-189">The image's format is specified as a [PictureFormat](/javascript/api/excel/excel.pictureformat) enum passed to `getAsImage`.</span></span>

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

## <a name="delete-shapes"></a><span data-ttu-id="5727f-190">Supprimer des formes</span><span class="sxs-lookup"><span data-stu-id="5727f-190">Delete shapes</span></span>

<span data-ttu-id="5727f-191">Les formes sont supprimées de la feuille de calcul à `Shape` l’aide de la méthode de `delete` l’objet.</span><span class="sxs-lookup"><span data-stu-id="5727f-191">Shapes are removed from the worksheet with the `Shape` object's `delete` method.</span></span> <span data-ttu-id="5727f-192">Aucune autre métadonnée n’est nécessaire.</span><span class="sxs-lookup"><span data-stu-id="5727f-192">No other metadata is needed.</span></span>

<span data-ttu-id="5727f-193">L’exemple de code suivant supprime toutes les formes de **MyWorksheet**.</span><span class="sxs-lookup"><span data-stu-id="5727f-193">The following code sample deletes all the shapes from **MyWorksheet**.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="5727f-194">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="5727f-194">See also</span></span>

- [<span data-ttu-id="5727f-195">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="5727f-195">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="5727f-196">Utiliser des graphiques à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="5727f-196">Work with charts using the Excel JavaScript API</span></span>](excel-add-ins-charts.md)
