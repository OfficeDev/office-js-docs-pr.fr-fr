---
title: Instructions relatives aux icônes de style monolignes pour les compléments Office
description: Obtenir des instructions sur l’utilisation des icônes d’icône de style monoligne dans les compléments Office.
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: 36142e79853a0fad47963255eb9517acd0810920
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44607693"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a>Instructions relatives aux icônes de style monolignes pour les compléments Office

Les iconographie de style monolignes sont utilisés dans Office 365. Si vous préférez que vos icônes correspondent au style frais sans abonnement Office 2013 +, reportez-vous à la rubrique [règles d’icône de style frais pour les compléments Office](add-in-icons-fresh.md).

## <a name="office-monoline-visual-style"></a>Style visuel monoligne Office

L’objectif du style monoligne est de disposer de iconographie cohérentes, claires et accessibles pour communiquer des actions et des fonctionnalités avec des éléments visuels simples, de s’assurer que les icônes sont accessibles à tous les utilisateurs et ont un style cohérent avec ceux utilisés ailleurs dans Windows.

Les recommandations suivantes sont destinées aux développeurs tiers qui souhaitent créer des icônes pour les fonctionnalités qui seront cohérentes avec les icônes qui présentent déjà des produits Office.

### <a name="design-principles"></a>Principes de conception

-   Simple, propre, clair.
-   Contenir uniquement les éléments nécessaires.
-   Style d’icône inspiré par Windows.
-   Accessible à tous les utilisateurs.

#### <a name="conveying-meaning"></a>Transport de sens

-   Utilisez des éléments descriptifs, tels qu’une page, pour représenter un document ou une enveloppe qui représente le courrier.
-   Utilisez le même élément pour représenter le même concept, c’est-à-dire que le courrier est toujours représenté par une enveloppe, pas par un cachet.
-   Utilisez une métaphore de base lors du développement du concept.

#### <a name="reduction-of-elements"></a>Réduction des éléments

-   Réduisez l’icône à sa signification fondamentale, en utilisant uniquement des éléments essentiels à la métaphore.
-   Limitez le nombre d’éléments d’une icône à deux, quelle que soit la taille des icônes.

#### <a name="consistency"></a>Concordance

Les tailles, l’organisation et la couleur des icônes doivent être cohérentes.

#### <a name="styling"></a>Style

##### <a name="perspective"></a>Perspective

Par défaut, les icônes monolignes sont dirigées vers l’avant. Certains éléments nécessitant une perspective et/ou une rotation, tels qu’un cube, sont autorisés, mais les exceptions doivent être réduites au minimum.

##### <a name="embellishment"></a>Ornement

La numérotation monoligne est un style minimal minimal. Tout utilise la couleur plate, ce qui signifie qu’il n’y a pas de dégradés, de textures ou de sources lumineuses.

## <a name="designing"></a>Créé

### <a name="sizes"></a>Quelle

Nous vous recommandons de créer chaque icône dans toutes ces tailles afin de prendre en charge les périphériques haute résolution. Les tailles absolument *requises* sont 16px, 20px et des, car il s’agit de la taille 100%.

**16px, 20px, des, des, 40px, 48px, 64px, 80px, 96px**

### <a name="layout"></a>Disposition

Voici un exemple de mise en page d’icône avec un modificateur.

![Exemple d’icône avec un modificateur](../images/monolineicon1.png)  ![Le même exemple avec une légende d’arrière-plan de grille pour la base, le modificateur, le remplissage et le découpage.](../images/monolineicon2.png)

#### <a name="elements"></a>Éléments

- **Base**: concept principal représenté par l’icône. Il s’agit généralement du seul visuel nécessaire pour l’icône, mais parfois le concept principal peut être amélioré avec un élément secondaire, un modificateur.

- **Modificateur** Tout élément qui chevauche la base ; autrement dit, un modificateur qui représente généralement une action ou un État. Il modifie l’élément base en agissant comme un ajout, une modification ou un descripteur.

![Grille contenant les zones de la zone de base et du modificateur.](../images/monolineicon3.png)

### <a name="construction"></a>Construction

#### <a name="element-placement"></a>Placement d’un élément

Les éléments de base sont placés au centre de l’icône dans le remplissage. Si elle ne peut pas être placée comme étant parfaitement centrée, la base doit se présenter sous la partie supérieure droite. Dans l’exemple suivant, l’icône est parfaitement centrée :

![Image montrant une icône parfaitement centré](../images/monolineicon4.png)

Dans l’exemple suivant, l’icône est erring vers la gauche.

![Image illustrant une icône errs vers la gauche](../images/monolineicon5.png)

Les modificateurs sont presque toujours placés dans le coin inférieur droit de la zone de dessin icône. Dans certains cas rares, les modificateurs sont placés dans un coin différent. Par exemple, si l’élément base ne peut pas être reconnaissable avec le modificateur dans le coin inférieur droit, envisagez de le placer dans le coin supérieur gauche.

![Image montrant quelques icônes avec le modificateur dans la partie inférieure droite, mais un autre avec le modificateur dans le coin supérieur gauche](../images/monolineicon6.png)

#### <a name="padding"></a>Remplissage

Chaque icône de taille possède une quantité spécifiée de remplissage entourant l’icône. L’élément base reste dans le remplissage, mais le modificateur doit se déplacer jusqu’au bord de la zone de dessin, en s’étendant en dehors du---de remplissage vers le bord de la bordure de l’icône. Les images suivantes montrent le remplissage recommandé à utiliser pour chaque taille d’icône.

|**16px**|**20px**|**24px**|**32px**|**40px**|**48px**|**64px**|**80 px**|**96px**|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![icône 16 PX](../images/monolineicon7.png)|![icône 20 px](../images/monolineicon8.png)|![icône 24 PX](../images/monolineicon9.png)|![icône 32 PX](../images/monolineicon10.png)|![icône 40 PX](../images/monolineicon11.png)|![icône 48 px](../images/monolineicon12.png)|![icône 64 px](../images/monolineicon13.png)|![icône 80 PX](../images/monolineicon14.png)|![icône 96 PX](../images/monolineicon15.png)|

#### <a name="line-weights"></a>Épaisseurs de trait

La numérotation monoligne est un style dominée par ligne et avec contour. En fonction de la taille que vous générez, l’icône doit utiliser les épaisseurs de trait suivantes.

|**Taille de l’icône :**|**16px**|**20px**|**24px**|**32px**|**40px**|**48px**|**64px**|**80 px**|**96px**|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|**Épaisseur de trait :**|1px|1px|1px|1px|2px|2px|2px|2px|3px|
||![icône 16 PX](../images/monolineicon16.png)|![icône 20 px](../images/monolineicon17.png)|![icône 24 PX](../images/monolineicon18.png)|![icône 32 PX](../images/monolineicon19.png)|![icône 40 PX](../images/monolineicon20.png)|![icône 48 px](../images/monolineicon21.png)|![icône 64 px](../images/monolineicon22.png)|![icône 80 PX](../images/monolineicon23.png)|![icône 96 PX](../images/monolineicon24.png)|

#### <a name="cutouts"></a>Découpages

Lorsqu’un élément Icon est placé au-dessus d’un autre élément, une découpe (de l’élément du bas) est utilisée pour fournir l’espace entre les deux éléments, principalement pour des raisons de lisibilité. Cela se produit généralement lorsqu’un modificateur est placé au-dessus d’un élément de base, mais il y a également des cas où aucun des éléments n’est un modificateur. Ces découpages entre les deux éléments sont parfois appelés « Gap ».

La taille de l’écart doit être de la même largeur que l’épaisseur de trait utilisée sur cette taille. Si vous créez une icône 16px, la largeur de l’intervalle est 1 pixel et s’il s’agit d’une icône 48px, l’intervalle doit être "Medium. L’exemple suivant montre une icône des avec un intervalle de 1 pixel entre le modificateur et la base sous-jacente.

![icône des avec un intervalle de 1 pixel entre le modificateur et la base sous-jacente](../images/monolineicon25.png)

Dans certains cas, l’écart peut être augmenté de 1/"Medium si le modificateur a une arête diagonale ou courbée et que l’intervalle standard ne fournit pas une séparation suffisante. Cela affectera probablement seulement les icônes avec 1 pixel épaisseur de trait ; 16px, 20px, des et des.

#### <a name="background-fills"></a>Remplissages d’arrière-plan

La plupart des icônes du jeu d’icônes monolignes nécessitent des remplissages d’arrière-plan. Toutefois, dans certains cas, l’objet n’aura pas de remplissage naturellement, aucun remplissage ne doit donc être appliqué. Les icônes suivantes ont un remplissage blanc :

![Cinq icônes ont un remplissage blanc](../images/monolineicon26.png)

Les icônes suivantes n’ont pas de remplissage. (L’icône d’engrenage est incluse pour indiquer que le trou central n’est pas rempli.) ![Cinq icônes sans remplissage](../images/monolineicon27.png)

##### <a name="best-practices-for-fills"></a>Meilleures pratiques pour les remplissages

###### <a name="dos"></a>Étendue

- Remplissez tous les éléments qui ont une limite définie, et qu’ils ont naturellement un remplissage.
- Utilisez une forme distincte pour créer le remplissage de l’arrière-plan.
- Utiliser le remplissage de l' **arrière-plan** de la [palette de couleurs](#color).
- Conservez la séparation des pixels entre les éléments qui se chevauchent.
- Remplissage entre plusieurs objets.

###### <a name="donts"></a>Don’t

- Ne remplissez pas les objets qui ne seraient pas naturellement remplis ; par exemple, un trombone.
- Ne pas remplir les crochets.
- Ne pas remplir les chiffres ou les caractères alpha.

### <a name="color"></a>Couleur

La palette de couleurs a été conçue pour des fins de simplicité et d’accessibilité. Elle contient 4 couleurs neutres et deux variantes pour le bleu, le vert, le jaune, le rouge et le violet. La couleur orange n’est intentionnellement pas incluse dans la palette de couleurs de l’icône monoligne. Chaque couleur est destinée à être utilisée de différentes manières, comme décrit dans cette section.

#### <a name="palette"></a>Texture

![Les quatre nuances de gris en monoligne](../images/monoline-grayshades.png)

![Palette de couleurs sur une seule ligne](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a>Utilisation de la couleur

Dans la palette de couleurs monolignes, toutes les couleurs ont des variantes autonomes, de contour et de remplissage. En règle générale, les éléments sont construits avec un remplissage et une bordure. Les couleurs sont appliquées dans l’un des modèles suivants :

- La couleur autonome uniquement pour les objets qui n’ont pas de remplissage.
- La bordure utilise la couleur de contour et le remplissage utilise la couleur de remplissage.
- La bordure utilise la couleur autonome et le remplissage utilise la couleur de remplissage de l’arrière-plan.

Voici des exemples d’utilisation de la couleur.

![Trois icônes avec une couleur dans une bordure ou un remplissage ou les deux](../images/monolineicon28.png)

La situation la plus courante consistera à utiliser un élément gris foncé avec remplissage d’arrière-plan.

Lorsque vous utilisez un remplissage coloré, il doit toujours être associé à la couleur de contour correspondante. Par exemple, le remplissage bleu ne doit être utilisé qu’avec un contour bleu. Toutefois, il existe deux exceptions à cette règle générale :

- Le remplissage d’arrière-plan peut être utilisé avec n’importe quelle couleur autonome.
- Le remplissage gris clair peut être utilisé avec deux couleurs de contour différentes : gris foncé ou gris moyen.

#### <a name="when-to-use-color"></a>Quand utiliser la couleur

La couleur doit être utilisée pour indiquer la signification de l’icône plutôt que pour l’ornement. Il doit **mettre en surbrillance l’action** pour l’utilisateur. Lorsqu’un modificateur est ajouté à un élément de base avec une couleur, l’élément de base est généralement transformé en gris foncé et remplissage d’arrière-plan de sorte que le modificateur puisse être l’élément de couleur, comme le cas ci-dessous avec le modificateur « X » ajouté à la base de l’image dans l’icône la plus à gauche du jeu suivant.

![Cinq icônes qui utilisent la couleur](../images/monolineicon29.png)

Vous devez limiter vos icônes à **une** couleur supplémentaire, autre que le contour et le remplissage mentionnés ci-dessus. Toutefois, il est possible d’utiliser davantage de couleurs s’il est vital pour sa métaphore, avec une limite de deux couleurs supplémentaires autres que le gris. Dans de rares cas, il existe des exceptions lorsque des couleurs supplémentaires sont nécessaires. Voici des exemples intéressants d’icônes qui n’utilisent qu’une seule couleur.

  ![Image de cinq icônes avec une couleur chacune](../images/monolineicon30.png)

Toutefois, les icônes suivantes utilisent trop de couleurs.

  ![Image de cinq icônes avec plusieurs couleurs](../images/monolineicon31.png)


Utiliser le **gris moyen** pour le « contenu » intérieur, comme le quadrillage dans une icône d’une feuille de calcul. Des couleurs d’intérieur supplémentaires sont utilisées lorsque le contenu doit afficher le comportement du contrôle.

![Cinq icônes avec des éléments intérieurs de gris moyen](../images/monolineicon32.png)

#### <a name="text-lines"></a>Lignes de texte

Lorsque les lignes de texte se trouvent dans un « conteneur » (par exemple, du texte dans un document), utilisez le gris moyen. Les lignes de texte qui ne se trouvent pas dans un conteneur doivent être **gris foncées**.

### <a name="text"></a>Texte

Évitez d’utiliser des caractères de texte dans les icônes. Étant donné que les produits Office sont utilisés dans le monde entier, nous souhaitons conserver les icônes aussi indépendantes que possible.

## <a name="production"></a>Production

### <a name="icon-file-format"></a>Format du fichier d’icônes

Les dernières icônes doivent être enregistrées en tant que fichiers image. png. Utilisez le format PNG avec un arrière-plan transparent et avez une profondeur de 32 bits.
