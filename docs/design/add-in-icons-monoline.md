---
title: Instructions relatives aux icônes de style monoligne pour Office de recherche
description: Recommandations en matière d’utilisation d’icônes de style Monoline dans Office des modules.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 0e8bf4f39ddbad457df7d033a08836825d9e1d3f
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938108"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a>Instructions relatives aux icônes de style monoligne pour Office de recherche

L’iconographie de style monoligne est utilisée dans Office applications. Si vous préférez que vos icônes correspondent au style Fresh de Office 2013+ sans abonnement, consultez les instructions relatives aux [icônes](add-in-icons-fresh.md)de style Fresh pour les Office.

## <a name="office-monoline-visual-style"></a>Office Style visuel monoligne

L’objectif du style Monoline est d’avoir une iconographie cohérente, claire et accessible pour communiquer des actions et des fonctionnalités avec des éléments visuels simples, garantir que les icônes sont accessibles à tous les utilisateurs et avoir un style cohérent avec ceux utilisés ailleurs dans Windows.

Les instructions suivantes sont pour les développeurs tiers qui souhaitent créer des icônes pour des fonctionnalités cohérentes avec les icônes déjà présentes Office produits.

### <a name="design-principles"></a>Principes de conception

- Simple, propre, clair.
- Contiennent uniquement les éléments nécessaires.
- Inspired by Windows icon style.
- Accessible à tous les utilisateurs.

#### <a name="convey-meaning"></a>Transmettre une signification

- Utilisez des éléments descriptifs tels qu’une page pour représenter un document ou une enveloppe pour représenter le courrier électronique.
- Utilisez le même élément pour représenter le même concept, c’est-à-dire que le courrier est toujours représenté par une enveloppe, et non par un cachet.
- Utilisez une métaphore principale pendant le développement de concepts.

#### <a name="reduction-of-elements"></a>Réduction des éléments

- Réduisez l’icône à sa signification principale, en utilisant uniquement les éléments essentiels à la métaphore.
- Limitez le nombre d’éléments d’une icône à deux, quelle que soit la taille de l’icône.

#### <a name="consistency"></a>Cohérence

Les tailles, la disposition et la couleur des icônes doivent être cohérentes.

#### <a name="styling"></a>Stylisme

##### <a name="perspective"></a>Perspective

Les icônes monolignes sont orientées vers l’avant par défaut. Certains éléments nécessitant une perspective et/ou une rotation, tels qu’un cube, sont autorisés, mais les exceptions doivent être conservées au minimum.

##### <a name="embellishment"></a>Enjolivement

Monoline est un style minimal. Tout utilise une couleur plate, ce qui signifie qu’il n’y a pas de dégradés, de textures ou de sources de lumière.

## <a name="designing"></a>Conception

### <a name="sizes"></a>Tailles

Nous vous recommandons de produire chaque icône de toutes ces tailles pour prendre en charge les appareils à hautes dimensions. Les tailles *absolument requises* sont de 16 px, 20 px et 32 px, car il s’s’il s’tt de la taille 100 %.

**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**

> [!IMPORTANT]
> Pour obtenir une image représentant l’icône représentant votre application, voir Créer des listes efficaces dans [AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) et dans Office pour la taille et d’autres exigences.

### <a name="layout"></a>Disposition

Voici un exemple de mise en page d’icône avec un modificateur.

![Diagramme de l’icône avec modificateur en bas à droite.](../images/monolineicon1.png)  ![Diagramme de la même icône avec un arrière-plan de grille et des légendes ajoutés pour la base, le modificateur, l’remplissage et le cutout.](../images/monolineicon2.png)

#### <a name="elements"></a>Éléments

- **Base**: concept principal représenté par l’icône. Il s’agit généralement du seul élément visuel nécessaire à l’icône, mais le concept principal peut parfois être amélioré avec un élément secondaire, un modificateur.

- **Modificateur** Tout élément qui superpose la base ; autrement dit, un modificateur qui représente généralement une action ou un état. Il modifie l’élément de base en agissant comme un ajout, une modification ou un descripteur.

![Diagramme de la grille avec zones de base et de modificateur appelées.](../images/monolineicon3.png)

### <a name="construction"></a>Construction

#### <a name="element-placement"></a>Placement des éléments

Les éléments de base sont placés au centre de l’icône dans le remplissage. Si elle ne peut pas être parfaitement centrée, la base doit se placer en haut à droite. Dans l’exemple suivant, l’icône est parfaitement centrée.

![Diagramme montrant une icône parfaitement centrée.](../images/monolineicon4.png)

Dans l’exemple suivant, l’icône se trouve à gauche.

![Diagramme montrant l’icône qui se place à gauche d'1 px.](../images/monolineicon5.png)

Les modificateurs sont presque toujours placés dans le coin inférieur droit de la zone de dessin de l’icône. Dans certains cas, les modificateurs sont placés dans un autre coin. Par exemple, si l’élément de base ne serait pas reconnu par le modificateur dans le coin inférieur droit, envisagez de le placer dans le coin supérieur gauche.

![Diagramme montrant quatre icônes avec le modificateur en bas à droite et une icône avec le modificateur dans le coin supérieur gauche.](../images/monolineicon6.png)

#### <a name="padding"></a>Remplissage

Chaque icône de taille possède une quantité spécifiée de remplissage autour de l’icône. L’élément de base reste dans l’espacement, mais le modificateur doit pointer jusqu’au bord de la zone de dessin, en s’étendant en dehors du remplissage jusqu’au bord de la bordure de l’icône. Les images suivantes montrent le remplissage recommandé à utiliser pour chacune des tailles d’icône.

|**16px**|**20px**|**24px**|**32px**|**40px**|**48px**|**64px**|**80 px**|**96 px**|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![Icône de 16 px avec remplissage de 0 px.](../images/monolineicon7.png)|![Icône de 20 px avec remplissage 1 px.](../images/monolineicon8.png)|![Icône de 24 px avec remplissage 1 px.](../images/monolineicon9.png)|![Icône de 32 px avec remplissage de 2 px.](../images/monolineicon10.png)|![Icône de 40 px avec remplissage de 2 px.](../images/monolineicon11.png)|![Icône de 48 px avec remplissage de 3 px.](../images/monolineicon12.png)|![Icône de 64 px avec remplissage 4 px.](../images/monolineicon13.png)|![Icône de 80 px avec remplissage de 5 px.](../images/monolineicon14.png)|![Icône de 96 px avec remplissage de 6 px.](../images/monolineicon15.png)|

#### <a name="line-weights"></a>Poids des lignes

Le monoligne est un style en courbes et en contours. Selon la taille que vous produisez, l’icône doit utiliser les poids de ligne suivants.

|Taille de l’icône :|16px|20px|24px|32px|40px|48px|64px|80 px|96 px|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|**Poids de ligne :**|1px|1px|1px|1px|2px|2px|2px|2px|3px|
|**Exemple d’icône :**|![Icône 16 px.](../images/monolineicon16.png)|![Icône 20 px.](../images/monolineicon17.png)|![Icône 24 px.](../images/monolineicon18.png)|![Icône 32 px.](../images/monolineicon19.png)|![Icône 40 px.](../images/monolineicon20.png)|![Icône 48 px.](../images/monolineicon21.png)|![Icône 64 px.](../images/monolineicon22.png)|![Icône 80 px.](../images/monolineicon23.png)|![Icône 96 px.](../images/monolineicon24.png)|

#### <a name="cutouts"></a>Cutouts

Lorsqu’un élément d’icône est placé au-dessus d’un autre élément, un cutout (de l’élément inférieur) est utilisé pour fournir de l’espace entre les deux éléments, principalement à des fins de lisibilité. Cela se produit généralement lorsqu’un modificateur est placé au-dessus d’un élément de base, mais dans certains cas, aucun des éléments n’est un modificateur. Ces coupures entre les deux éléments sont parfois appelées « intervalle ».

La taille de l’intervalle doit être identique à la largeur de trait utilisée sur cette taille. Si vous faites une icône de 16 px, la largeur de l’intervalle est de 1 px et s’il s’agit d’une icône de 48 px, l’écart doit être de 2 px. L’exemple suivant montre une icône de 32 px avec un intervalle de 1 px entre le modificateur et la base sous-jacente.

![Icône de 32 px avec un intervalle de 1 px entre le modificateur et la base sous-jacente.](../images/monolineicon25.png)

Dans certains cas, l’écart peut être augmenté de 1/2 px si le modificateur possède un bord diagonal ou courbé et que l’intervalle standard ne fournit pas une séparation suffisante. Cela affectera probablement uniquement les icônes avec une pondération de trait de 1 px : 16 px, 20 px, 24 px et 32 px.

#### <a name="background-fills"></a>Remplissages d’arrière-plan

La plupart des icônes du jeu d’icônes Monoline nécessitent des remplissages d’arrière-plan. Toutefois, dans certains cas, l’objet n’a pas naturellement de remplissage, aucun remplissage ne doit donc être appliqué. Les icônes suivantes ont un remplissage blanc.

![Compilation de cinq icônes avec remplissage blanc.](../images/monolineicon26.png)

Les icônes suivantes n’ont pas de remplissage. (L’icône d’engrenage est incluse pour montrer que le centre du centre n’est pas rempli.)

![Compilation de cinq icônes sans remplissage.](../images/monolineicon27.png)

##### <a name="best-practices-for-fills"></a>Meilleures pratiques en matière de remplissages

###### <a name="dos"></a>À faire

- Remplissez tout élément qui a une limite définie et qui aurait naturellement un remplissage.
- Utilisez une forme distincte pour créer le remplissage d’arrière-plan.
- Utilisez le **remplissage d’arrière-plan** à partir de [la palette de couleurs.](#color)
- Conservez la séparation des pixels entre les éléments qui se chevauchent.
- Remplissage entre plusieurs objets.

###### <a name="donts"></a>À ne pas faire

- Ne remplissez pas les objets qui ne seraient pas naturellement remplis ; par exemple, un paperclip.
- Ne pas remplir les crochets.
- Ne pas remplir les chiffres ou les caractères alpha.

### <a name="color"></a>Couleur

La palette de couleurs a été conçue pour simplifier et accessibilité. Il contient 4 couleurs neutres et deux variantes pour le bleu, le vert, le jaune, le rouge et le violet. L’orange n’est intentionnellement pas inclus dans la palette de couleurs de l’icône Monoline. Chaque couleur est destinée à être utilisée de manière spécifique, comme indiqué dans cette section.

#### <a name="palette"></a>Palette

![Les quatre nuances de gris en monoligne : gris foncé pour un contour ou autonome, gris moyen pour le plan ou le contenu, gris très clair pour le remplissage d’arrière-plan et gris clair pour le remplissage.](../images/monoline-grayshades.png)

![La palette de couleurs en monoligne inclut une nuance de bleu, vert, jaune, rouge et violet pour les lignes autonomes, les contours et le remplissage.](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a>Comment utiliser la couleur

Dans la palette de couleurs Monoline, toutes les couleurs ont des variantes Autonome, Plan et Remplissage. En règle générale, les éléments sont construits avec un remplissage et une bordure. Les couleurs sont appliquées dans l’un des motifs suivants.

- Couleur autonome uniquement pour les objets sans remplissage.
- La bordure utilise la couleur Plan et le remplissage utilise la couleur Remplissage.
- La bordure utilise la couleur autonome et le remplissage utilise la couleur de remplissage d’arrière-plan.

Voici des exemples d’utilisation de couleur.

![Compilation de trois icônes avec une couleur dans une bordure ou un remplissage ou les deux.](../images/monolineicon28.png)

La situation la plus courante est qu’un élément utilise le gris foncé autonome avec remplissage d’arrière-plan.

Lors de l’utilisation d’un remplissage coloré, il doit toujours être avec sa couleur de plan correspondante. Par exemple, le remplissage bleu ne doit être utilisé qu’avec le contour bleu. Mais il existe deux exceptions à cette règle générale.

- Le remplissage d’arrière-plan peut être utilisé avec n’importe quelle couleur autonome.
- Le remplissage gris clair peut être utilisé avec deux couleurs plan différentes : gris foncé ou gris moyen.

#### <a name="when-to-use-color"></a>Quand utiliser la couleur

La couleur doit être utilisée pour transmettre la signification de l’icône plutôt que pour l’enjolivement. **L’action doit être mise en surbrillant** pour l’utilisateur. Lorsqu’un modificateur est ajouté à un élément de base qui a une couleur, l’élément de base est généralement transformé en gris foncé et remplissage d’arrière-plan afin que le modificateur puisse être l’élément de couleur, comme le cas ci-dessous avec le modificateur « X » ajouté à la base d’image dans l’icône la plus à gauche du jeu suivant.

![Compilation de cinq icônes qui utilisent la couleur.](../images/monolineicon29.png)

Vous devez limiter vos icônes à **une** couleur supplémentaire, autre que le plan et le remplissage mentionnés ci-dessus. Toutefois, il est possible d’utiliser davantage de couleurs s’il est essentiel pour sa métaphore, avec une limite de deux couleurs supplémentaires autres que le gris. Dans de rares cas, il existe des exceptions lorsque d’autres couleurs sont nécessaires. Voici de bons exemples d’icônes qui utilisent une seule couleur.

  ![Compilation de cinq icônes qui utilisent chacune une couleur.](../images/monolineicon30.png)

Mais les icônes suivantes utilisent trop de couleurs.

  ![Compilation de cinq icônes qui utilisent chacune plusieurs couleurs.](../images/monolineicon31.png)

Utilisez **un gris moyen** pour le « contenu » intérieur, tel que les lignes de grille dans une icône d’une feuille de calcul. Des couleurs intérieures supplémentaires sont utilisées lorsque le contenu doit afficher le comportement du contrôle.

![Compilation de cinq icônes avec des éléments intérieurs gris moyen.](../images/monolineicon32.png)

#### <a name="text-lines"></a>Lignes de texte

Lorsque des lignes de texte sont dans un « conteneur » (par exemple, du texte sur un document), utilisez un gris moyen. Les lignes de texte qui ne sont pas dans un conteneur doivent être **en gris foncé.**

### <a name="text"></a>Texte

Évitez d’utiliser des caractères de texte dans les icônes. Étant donné Office produits sont utilisés dans le monde entier, nous voulons conserver les icônes aussi neutres que possible en langage.

## <a name="production"></a>Production

### <a name="icon-file-format"></a>Format de fichier d’icône

Les icônes finales doivent être enregistrées sous forme .png fichiers image. Utilisez le format PNG avec un arrière-plan transparent et une profondeur 32 bits.

## <a name="see-also"></a>Voir aussi

- [Élément de manifeste d’icône](../reference/manifest/icon.md)
- [Élément manifeste IconUrl](../reference/manifest/iconurl.md)
- [Élément manifeste HighResolutionIconUrl](../reference/manifest/highresolutioniconurl.md)
- [Créer une icône pour votre add-in](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in)
