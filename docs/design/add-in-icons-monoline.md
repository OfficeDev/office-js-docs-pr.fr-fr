---
title: Instructions relatives aux icônes de style monoligne pour les compléments Office
description: Instructions pour l’utilisation d’icônes de style Monoline dans les compléments Office.
ms.date: 03/30/2021
ms.localizationpriority: medium
ms.openlocfilehash: 7af7cbb7539ee2ae27efcadd4739f926cc81547a
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467068"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a>Instructions relatives aux icônes de style monoligne pour les compléments Office

L’iconographie de style monoligne est utilisée dans les applications Office. Si vous préférez que vos icônes correspondent au style Fresh d’Office 2013+, consultez [les instructions relatives aux icônes de style Frais pour les compléments Office](add-in-icons-fresh.md).

## <a name="office-monoline-visual-style"></a>Style visuel Monoline Office

L’objectif du style Monoline est d’avoir une syntaxe cohérente, claire et accessible pour communiquer l’action et les fonctionnalités avec des visuels simples, s’assurer que les icônes sont accessibles à tous les utilisateurs et avoir un style cohérent avec ceux utilisés ailleurs dans Windows.

Les instructions suivantes s’appliquent aux développeurs tiers qui souhaitent créer des icônes pour les fonctionnalités qui seront cohérentes avec les icônes déjà présentes dans les produits Office.

### <a name="design-principles"></a>Principes de conception

- Simple, propre, clair.
- Contiennent uniquement les éléments nécessaires.
- Inspiré du style d’icône Windows.
- Accessible à tous les utilisateurs.

#### <a name="convey-meaning"></a>Transmettre la signification

- Utilisez des éléments descriptifs tels qu’une page pour représenter un document ou une enveloppe pour représenter le courrier.
- Utilisez le même élément pour représenter le même concept, c’est-à-dire que le courrier est toujours représenté par une enveloppe, et non par un tampon.
- Utilisez une métaphore principale lors du développement de concept.

#### <a name="reduction-of-elements"></a>Réduction des éléments

- Réduisez l’icône à sa signification principale, en utilisant uniquement les éléments essentiels à la métaphore.
- Limitez le nombre d’éléments d’une icône à deux, quelle que soit la taille de l’icône.

#### <a name="consistency"></a>Cohérence

Les tailles, la disposition et la couleur des icônes doivent être cohérentes.

#### <a name="styling"></a>Style

##### <a name="perspective"></a>Perspective

Les icônes monolignes sont orientées vers l’avant par défaut. Certains éléments qui nécessitent une perspective et/ou une rotation, tels qu’un cube, sont autorisés, mais les exceptions doivent être conservées au minimum.

##### <a name="embellishment"></a>Embellissement

Monoline est un style minimal propre. Tout utilise une couleur plate, ce qui signifie qu’il n’y a pas de dégradés, de textures ou de sources de lumière.

## <a name="designing"></a>Conception

### <a name="sizes"></a>Tailles

Nous vous recommandons de produire chaque icône de toutes ces tailles pour prendre en charge les appareils haute résolution. Les tailles absolument *requises* sont 16 px, 20 px et 32 px, car il s’agit des tailles de 100 %.

**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**

> [!IMPORTANT]
> Pour obtenir une image représentant l’icône représentative de votre complément, consultez [Créer des descriptions efficaces dans AppSource et dans Office](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) pour connaître la taille et d’autres exigences.

### <a name="layout"></a>Disposition

Voici un exemple de disposition d’icône avec un modificateur.

![Diagramme de l’icône avec modificateur en bas à droite.](../images/monolineicon1.png)  ![Diagramme de la même icône avec l’arrière-plan et les légendes de grille ajoutés pour la base, le modificateur, le remplissage et le découpage.](../images/monolineicon2.png)

#### <a name="elements"></a>Éléments

- **Base** : concept principal représenté par l’icône. Il s’agit généralement du seul visuel nécessaire pour l’icône, mais parfois le concept principal peut être amélioré avec un élément secondaire, un modificateur.

- **Modificateur** Tout élément qui superpose la base ; autrement dit, un modificateur qui représente généralement une action ou un état. Il modifie l’élément de base en agissant comme un ajout, une modification ou un descripteur.

![Diagramme de la grille avec les zones de base et de modification indiquées.](../images/monolineicon3.png)

### <a name="construction"></a>Construction

#### <a name="element-placement"></a>Placement d’élément

Les éléments de base sont placés au centre de l’icône dans le remplissage. S’il ne peut pas être placé parfaitement centré, la base doit errer en haut à droite. Dans l’exemple suivant, l’icône est parfaitement centrée.

![Diagramme montrant l’icône parfaitement centrée.](../images/monolineicon4.png)

Dans l’exemple suivant, l’icône est erronée à gauche.

![Diagramme montrant l’icône qui se trompe à gauche par 1 px.](../images/monolineicon5.png)

Les modificateurs sont presque toujours placés dans le coin inférieur droit du canevas d’icône. Dans certains cas rares, les modificateurs sont placés dans un autre coin. Par exemple, si l’élément de base est méconnaissable avec le modificateur dans le coin inférieur droit, envisagez de le placer dans le coin supérieur gauche.

![Diagramme montrant quatre icônes avec le modificateur en bas à droite et une icône avec le modificateur en haut à gauche.](../images/monolineicon6.png)

#### <a name="padding"></a>Rembourrage

Chaque icône de taille a une quantité spécifiée de remplissage autour de l’icône. L’élément de base reste dans le remplissage, mais le modificateur doit buter jusqu’au bord du canevas, en s’étendant en dehors du remplissage jusqu’au bord de la bordure de l’icône. Les images suivantes montrent le remplissage recommandé à utiliser pour chacune des tailles d’icône.

|**16px**|**20px**|**24px**|**32px**|**40px**|**48px**|**64px**|**80 px**|**96 px**|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![Icône 16 px avec remplissage 0px.](../images/monolineicon7.png)|![Icône 20 px avec remplissage 1px.](../images/monolineicon8.png)|![Icône 24 px avec remplissage 1px.](../images/monolineicon9.png)|![Icône 32 px avec remplissage 2px.](../images/monolineicon10.png)|![Icône 40 px avec remplissage 2px.](../images/monolineicon11.png)|![Icône 48 px avec remplissage 3px.](../images/monolineicon12.png)|![Icône 64 px avec remplissage 4px.](../images/monolineicon13.png)|![Icône 80 px avec remplissage de 5 pixels.](../images/monolineicon14.png)|![Icône 96 px avec remplissage 6px.](../images/monolineicon15.png)|

#### <a name="line-weights"></a>Épaisseurs de ligne

Monoline est un style dominé par des lignes et des formes hiérarchiques. Selon la taille que vous produisez, l’icône doit utiliser les poids de ligne suivants.

|Taille de l’icône :|16px|20px|24px|32px|40px|48px|64px|80 px|96 px|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|**Épaisseur de ligne :**|1px|1px|1px|1px|2px|2px|2px|2px|3px|
|**Exemple d’icône :**|![Icône 16 px.](../images/monolineicon16.png)|![Icône 20 px.](../images/monolineicon17.png)|![Icône 24 px.](../images/monolineicon18.png)|![Icône 32 px.](../images/monolineicon19.png)|![Icône 40 px.](../images/monolineicon20.png)|![Icône 48 px.](../images/monolineicon21.png)|![Icône 64 px.](../images/monolineicon22.png)|![Icône 80 px.](../images/monolineicon23.png)|![Icône 96 px.](../images/monolineicon24.png)|

#### <a name="cutouts"></a>Découpes

Lorsqu’un élément d’icône est placé au-dessus d’un autre élément, un découpage (de l’élément inférieur) est utilisé pour fournir de l’espace entre les deux éléments, principalement à des fins de lisibilité. Cela se produit généralement lorsqu’un modificateur est placé au-dessus d’un élément de base, mais il existe également des cas où aucun des éléments n’est un modificateur. Ces découpages entre les deux éléments sont parfois appelés « écart ».

La taille de l’écart doit être de la même largeur que le poids de trait utilisé sur cette taille. Si vous créez une icône de 16 px, la largeur de l’espace est de 1 px et, s’il s’agit d’une icône de 48 px, l’écart doit être de 2 px. L’exemple suivant montre une icône de 32 px avec un écart de 1 px entre le modificateur et la base sous-jacente.

![Icône 32 px avec un intervalle de 1 px entre le modificateur et la base sous-jacente.](../images/monolineicon25.png)

Dans certains cas, l’écart peut être augmenté de 1/2 px si le modificateur a un bord diagonal ou courbé et que l’écart standard ne fournit pas suffisamment de séparation. Cela n’affectera probablement que les icônes avec un poids de ligne de 1 px : 16 px, 20 px, 24 px et 32 px.

#### <a name="background-fills"></a>Remplissages en arrière-plan

La plupart des icônes du jeu d’icônes Monoline nécessitent des remplissages d’arrière-plan. Toutefois, dans certains cas, l’objet n’a pas naturellement de remplissage, donc aucun remplissage ne doit être appliqué. Les icônes suivantes ont un remplissage blanc.

![Compilation de cinq icônes avec remplissage blanc.](../images/monolineicon26.png)

Les icônes suivantes n’ont pas de remplissage. (L’icône d’engrenage est incluse pour montrer que le trou central n’est pas rempli.)

![Compilation de cinq icônes sans remplissage.](../images/monolineicon27.png)

##### <a name="best-practices-for-fills"></a>Meilleures pratiques pour les remplissages

###### <a name="dos"></a>À faire

- Remplissez tout élément qui a une limite définie et qui aurait naturellement un remplissage.
- Utilisez une forme distincte pour créer le remplissage d’arrière-plan.
- Utilisez le **remplissage d’arrière-plan** à partir de la [palette de couleurs](#color).
- Conservez la séparation des pixels entre les éléments qui se chevauchent.
- Remplir entre plusieurs objets.

###### <a name="donts"></a>À ne pas faire

- Ne remplissez pas les objets qui ne seraient pas naturellement remplis ; par exemple, un paperclip.
- Ne remplissez pas les crochets.
- Ne remplissez pas derrière des nombres ou des caractères alpha.

### <a name="color"></a>Couleur

La palette de couleurs a été conçue pour la simplicité et l’accessibilité. Il contient 4 couleurs neutres et deux variantes pour le bleu, le vert, le jaune, le rouge et le violet. L’orange n’est intentionnellement pas inclus dans la palette de couleurs d’icône Monoline. Chaque couleur est destinée à être utilisée de manière spécifique, comme indiqué dans cette section.

#### <a name="palette"></a>Palette

![Les quatre nuances de gris en monoligne : gris foncé pour autonome ou contour, gris moyen pour le contour ou le contenu, gris très clair pour le remplissage d’arrière-plan et gris clair pour le remplissage.](../images/monoline-grayshades.png)

![La palette de couleurs en monoligne comprend une nuance de bleu, de vert, de jaune, de rouge et de violet pour l’autonome, le contour et le remplissage.](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a>Comment utiliser la couleur

Dans la palette de couleurs Monoline, toutes les couleurs ont des variantes Autonomes, Contour et Remplissage. En règle générale, les éléments sont construits avec un remplissage et une bordure. Les couleurs sont appliquées dans l’un des modèles suivants.

- Couleur autonome uniquement pour les objets qui n’ont pas de remplissage.
- La bordure utilise la couleur Contour et le remplissage utilise la couleur de remplissage.
- La bordure utilise la couleur autonome et le remplissage utilise la couleur de remplissage d’arrière-plan.

Voici des exemples d’utilisation de la couleur.

![Compilation de trois icônes avec une couleur dans une bordure ou un remplissage ou les deux.](../images/monolineicon28.png)

La situation la plus courante consiste à utiliser un élément en mode autonome gris foncé avec remplissage d’arrière-plan.

Lors de l’utilisation d’un remplissage coloré, il doit toujours être avec sa couleur de contour correspondante. Par exemple, le remplissage bleu ne doit être utilisé qu’avec le contour bleu. Mais il existe deux exceptions à cette règle générale.

- Le remplissage d’arrière-plan peut être utilisé avec n’importe quelle couleur autonome.
- Le remplissage gris clair peut être utilisé avec deux couleurs de contour différentes : gris foncé ou gris moyen.

#### <a name="when-to-use-color"></a>Quand utiliser la couleur

La couleur doit être utilisée pour transmettre la signification de l’icône plutôt que pour l’embellissement. Il doit **mettre en surbrillance l’action** pour l’utilisateur. Lorsqu’un modificateur est ajouté à un élément de base qui a de la couleur, l’élément de base est généralement transformé en gris foncé et en remplissage d’arrière-plan afin que le modificateur puisse être l’élément de couleur, comme le cas ci-dessous avec le modificateur « X » ajouté à la base de l’image dans l’icône la plus à gauche du jeu suivant.

![Compilation de cinq icônes qui utilisent la couleur.](../images/monolineicon29.png)

Vous devez limiter vos icônes à **une** couleur supplémentaire, autre que le contour et le remplissage mentionnés ci-dessus. Toutefois, plus de couleurs peuvent être utilisées si elle est essentielle pour sa métaphore, avec une limite de deux couleurs supplémentaires autres que le gris. Dans de rares cas, il existe des exceptions lorsque davantage de couleurs sont nécessaires. Voici de bons exemples d’icônes qui utilisent une seule couleur.

  ![Compilation de cinq icônes qui utilisent chacune une couleur.](../images/monolineicon30.png)

Mais les icônes suivantes utilisent trop de couleurs.

  ![Compilation de cinq icônes qui utilisent chacune plusieurs couleurs.](../images/monolineicon31.png)

Utilisez **le gris moyen** pour le « contenu » intérieur, tel que les lignes de grille dans une icône d’une feuille de calcul. Des couleurs intérieures supplémentaires sont utilisées lorsque le contenu doit afficher le comportement du contrôle.

![Compilation de cinq icônes avec des éléments intérieurs gris moyen.](../images/monolineicon32.png)

#### <a name="text-lines"></a>Lignes de texte

Lorsque les lignes de texte se trouvent dans un « conteneur » (par exemple, du texte sur un document), utilisez un gris moyen. Les lignes de texte qui ne sont pas dans un conteneur doivent être **gris foncé**.

### <a name="text"></a>Text

Évitez d’utiliser des caractères de texte dans les icônes. Étant donné que les produits Office sont utilisés dans le monde entier, nous voulons garder les icônes aussi neutres que possible.

## <a name="production"></a>Production

### <a name="icon-file-format"></a>Format de fichier d’icône

Les icônes finales doivent être enregistrées en tant que fichiers image .png. Utilisez le format PNG avec un arrière-plan transparent et une profondeur de 32 bits.

## <a name="see-also"></a>Voir aussi

- [Élément de manifeste d’icône](/javascript/api/manifest/icon)
- [Élément de manifeste IconUrl](/javascript/api/manifest/iconurl)
- [Élément manifeste HighResolutionIconUrl](/javascript/api/manifest/highresolutioniconurl)
- [Créer une icône pour votre complément](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in)
