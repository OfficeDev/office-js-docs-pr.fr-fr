---
title: PowerPoint d’aperçu JavaScript
description: Détails sur les API JavaScript PowerPoint à venir.
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: medium
ms.openlocfilehash: 406808b4b4ff16df72d9c37468696525c8be642f
ms.sourcegitcommit: e44a8109d9323aea42ace643e11717fb49f40baa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/15/2021
ms.locfileid: "61513990"
---
# <a name="powerpoint-javascript-preview-apis"></a>PowerPoint d’aperçu JavaScript

Les nouvelles API JavaScript PowerPoint sont d’abord introduites dans « aperçu », puis font partie d’un ensemble de conditions requises numérotées spécifiques une fois que des tests suffisants ont eu lieu et que les commentaires des utilisateurs ont été acquis.

Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Gestion des diapositives | Ajoute la prise en charge de l’ajout de diapositives, ainsi que de la gestion des mises en page des diapositives et des formes de base des diapositives. | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| Formes | Ajoute la prise en charge de l’obtention de références aux formes dans une diapositive. | [Shape](/javascript/api/powerpoint/powerpoint.shape) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les POWERPOINT JavaScript actuellement en prévisualisation. Pour obtenir la liste complète de toutes les API JavaScript PowerPoint (y compris les API de prévisualisation et les API publiées précédemment), voir toutes Excel API [JavaScript.](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)

| Classe | Champs | Description |
|:---|:---|:---|
|[BulletFormat (objet)](/javascript/api/powerpoint/powerpoint.bulletformat)|[visible](/javascript/api/powerpoint/powerpoint.bulletformat#visible)|Spécifie si les puces du paragraphe sont visibles.|
|[ParagraphFormat](/javascript/api/powerpoint/powerpoint.paragraphformat)|[bulletFormat](/javascript/api/powerpoint/powerpoint.paragraphformat#bulletFormat)|Représente le format de puce du paragraphe.|
||[horizontalAlignment](/javascript/api/powerpoint/powerpoint.paragraphformat#horizontalAlignment)|Représente l’alignement horizontal du paragraphe.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[fill](/javascript/api/powerpoint/powerpoint.shape#fill)|Renvoie la mise en forme de remplissage de cette forme.|
||[height](/javascript/api/powerpoint/powerpoint.shape#height)|Spécifie la hauteur, en points, de la forme.|
||[left](/javascript/api/powerpoint/powerpoint.shape#left)|Distance, en points, entre le côté gauche de la forme et le côté gauche de la diapositive.|
||[lineFormat](/javascript/api/powerpoint/powerpoint.shape#lineFormat)|Renvoie la mise en forme de ligne de cette forme.|
||[name](/javascript/api/powerpoint/powerpoint.shape#name)|Spécifie le nom de cette forme.|
||[textFrame](/javascript/api/powerpoint/powerpoint.shape#textFrame)|Renvoie l’objet textFrame d’une forme.|
||[top](/javascript/api/powerpoint/powerpoint.shape#top)|Distance, en points, entre le bord supérieur de la forme et le bord supérieur de la diapositive.|
||[type](/javascript/api/powerpoint/powerpoint.shape#type)|Renvoie le type de cette forme.|
||[width](/javascript/api/powerpoint/powerpoint.shape#width)|Spécifie la largeur, en points, de la forme.|
|[ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions)|[height](/javascript/api/powerpoint/powerpoint.shapeaddoptions#height)|Spécifie la hauteur, en points, de la forme.|
||[left](/javascript/api/powerpoint/powerpoint.shapeaddoptions#left)|Spécifie la distance, en points, entre le côté gauche de la forme et le côté gauche de la diapositive.|
||[top](/javascript/api/powerpoint/powerpoint.shapeaddoptions#top)|Spécifie la distance, en points, entre le bord supérieur de la forme et le bord supérieur de la diapositive.|
||[width](/javascript/api/powerpoint/powerpoint.shapeaddoptions#width)|Spécifie la largeur, en points, de la forme.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[addGeometricShape(geometricShapeType: PowerPoint. GeometricShapeType, options ? : PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addGeometricShape_geometricShapeType__options_)|Ajoute une forme géométrique à la diapositive.|
||[addLine(connectorType?: PowerPoint. ConnectorType, options ? : PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addLine_connectorType__options_)|Ajoute une ligne à la diapositive.|
||[addTextBox(text: string, options?: PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addTextBox_text__options_)|Ajoute une zone de texte à la diapositive avec le texte fourni comme contenu.|
|[ShapeFill](/javascript/api/powerpoint/powerpoint.shapefill)|[clear()](/javascript/api/powerpoint/powerpoint.shapefill#clear__)|Renvoie la mise en forme de remplissage de cette forme.|
||[foregroundColor](/javascript/api/powerpoint/powerpoint.shapefill#foregroundColor)|Représente la couleur de premier plan de remplissage de la forme au format HTML, au format #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[setSolidColor(color: string)](/javascript/api/powerpoint/powerpoint.shapefill#setSolidColor_color_)|Définit le format de remplissage d’un élément de graphique sur une couleur unie.|
||[Transparency](/javascript/api/powerpoint/powerpoint.shapefill#transparency)|Spécifie le pourcentage de transparence du remplissage sous la forme d’une valeur entre 0.0 (opaque) et 1.0 (clair).|
||[type](/javascript/api/powerpoint/powerpoint.shapefill#type)|Renvoie le type de remplissage de la forme.|
|[ShapeFont](/javascript/api/powerpoint/powerpoint.shapefont)|[bold](/javascript/api/powerpoint/powerpoint.shapefont#bold)|Représente le format de police Gras.|
||[color](/javascript/api/powerpoint/powerpoint.shapefont#color)|Représentation de code couleur HTML de la couleur du texte (par exemple, « #FF0000 » représente le rouge).|
||[italic](/javascript/api/powerpoint/powerpoint.shapefont#italic)|Représente le format de police Italique.|
||[name](/javascript/api/powerpoint/powerpoint.shapefont#name)|Représente le nom de la police (par exemple, « Calibri »).|
||[size](/javascript/api/powerpoint/powerpoint.shapefont#size)|Représente la taille de police en points (par exemple, 11).|
||[underline](/javascript/api/powerpoint/powerpoint.shapefont#underline)|Type de soulignement appliqué à la police.|
|[ShapeLineFormat](/javascript/api/powerpoint/powerpoint.shapelineformat)|[color](/javascript/api/powerpoint/powerpoint.shapelineformat#color)|Représente la couleur de trait au format HTML, au format #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[dashStyle](/javascript/api/powerpoint/powerpoint.shapelineformat#dashStyle)|Représente le style de tiret de la ligne.|
||[style](/javascript/api/powerpoint/powerpoint.shapelineformat#style)|Représente le style de trait de la forme.|
||[Transparency](/javascript/api/powerpoint/powerpoint.shapelineformat#transparency)|Spécifie le pourcentage de transparence de la ligne sous la forme d’une valeur entre 0.0 (opaque) et 1.0 (clair).|
||[visible](/javascript/api/powerpoint/powerpoint.shapelineformat#visible)|Spécifie si la mise en forme de trait d’un élément de forme est visible.|
||[weight](/javascript/api/powerpoint/powerpoint.shapelineformat#weight)|Représente l’épaisseur de ligne, en points.|
|[TextFrame](/javascript/api/powerpoint/powerpoint.textframe)|[autoSizeSetting](/javascript/api/powerpoint/powerpoint.textframe#autoSizeSetting)|Paramètres de resserrage automatique pour le cadre de texte.|
||[bottomMargin](/javascript/api/powerpoint/powerpoint.textframe#bottomMargin)|Représente la marge bas, en points du cadre du texte.|
||[deleteText()](/javascript/api/powerpoint/powerpoint.textframe#deleteText__)|Supprime tout le texte dans la textframe.|
||[hasText](/javascript/api/powerpoint/powerpoint.textframe#hasText)|Spécifie si le cadre de texte contient du texte.|
||[leftMargin](/javascript/api/powerpoint/powerpoint.textframe#leftMargin)|Représente la marge gauche, en points du cadre du texte.|
||[rightMargin](/javascript/api/powerpoint/powerpoint.textframe#rightMargin)|Représente la marge droite, en points du cadre du texte.|
||[textRange](/javascript/api/powerpoint/powerpoint.textframe#textRange)|Représente le texte lié à une forme, en plus des propriétés et des méthodes de manipulation du texte.|
||[topMargin](/javascript/api/powerpoint/powerpoint.textframe#topMargin)|Représente la marge du haut, en points du cadre du texte.|
||[verticalAlignment](/javascript/api/powerpoint/powerpoint.textframe#verticalAlignment)|Représente l’alignement vertical pour le style.|
||[wordWrap](/javascript/api/powerpoint/powerpoint.textframe#wordWrap)|Détermine si les lignes s’arrêtent automatiquement pour ajuster le texte à l’intérieur de la forme.|
|[TextRange](/javascript/api/powerpoint/powerpoint.textrange)|[police](/javascript/api/powerpoint/powerpoint.textrange#font)|Renvoie un `ShapeFont` objet qui représente les attributs de police de la plage de texte.|
||[getSubstring(start: number, length?: number)](/javascript/api/powerpoint/powerpoint.textrange#getSubstring_start__length_)|Renvoie un `TextRange` objet pour la sous-chaîne de la plage donnée.|
||[paragraphFormat](/javascript/api/powerpoint/powerpoint.textrange#paragraphFormat)|Représente le format de paragraphe de la plage de texte.|
||[text](/javascript/api/powerpoint/powerpoint.textrange#text)|Représente le contenu de texte brut de la plage de texte.|

## <a name="see-also"></a>Voir aussi

- [PowerPoint documentation de référence de l’API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour PowerPoint](powerpoint-api-requirement-sets.md)