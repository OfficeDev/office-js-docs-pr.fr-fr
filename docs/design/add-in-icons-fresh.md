---
title: Instructions relatives aux icônes de style frais pour les compléments Office
description: Instructions pour l’utilisation d’icônes de style frais dans les compléments Office.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: bd2cb372b79bef7f8c81deb778862f6bfd91d742
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467005"
---
# <a name="fresh-style-icon-guidelines-for-office-add-ins"></a>Instructions relatives aux icônes de style frais pour les compléments Office

Les versions Office 2013+ (perpétuelles) d’Office utilisent l’iconographie de style Fresh de Microsoft. Si vous préférez que vos icônes correspondent au style Monoline de Microsoft 365, consultez [les instructions relatives aux icônes de style Monoline pour les compléments Office](add-in-icons-monoline.md).

## <a name="office-fresh-visual-style"></a>Style visuel Office Fresh

Les icônes Fresh incluent uniquement des éléments communicatifs essentiels. Les éléments non essentiels, tels que la source de lumière, les dégradés et les perspectives, sont supprimés. Les icônes simplifiées prennent en charge l’analyse rapide des commandes et des contrôles. Suivez ce style pour s’adapter au mieux aux clients perpétuels Office.

## <a name="best-practices"></a>Meilleures pratiques

Suivez ces instructions lorsque vous créez vos icônes.

|À faire|À ne pas faire|
|:---|:---|
|Gardez les visuels simples et clairs, en vous concentrant sur les éléments clés de la communication.| N’utilisez pas d’artefacts qui rendent votre icône désordonnée.|
|Utilisez le langage d’icône Office pour représenter des comportements ou des concepts.|Ne réutilisez pas les glyphes Fabric Core pour les commandes de complément dans le ruban de l’application Office ou les menus contextuels. Les icônes Fabric Core sont différentes sur le plan stylistique et ne correspondent pas.|
|Réutilisez les métaphores visuelles d’Office courantes telles que le pinceau pour mettre en forme ou la loupe pour rechercher.|Ne réutilisez pas les métaphores visuelles pour différentes commandes. L’utilisation de la même icône pour différents comportements et concepts peut semer la confusion. |
|Redessinez vos icônes pour les réduire ou les agrandir. Prenez le temps de redessiner les découpages, les coins et des bords arrondis pour optimiser la netteté de ligne. |Ne redimensionnez pas vos icônes en réduisant ou en agrandissant leurs tailles. Cela peut entraîner une mauvaise qualité visuelle et des actions peu claires. Les icônes complexes créées dans une plus grande taille risquent de perdre en clarté si elles sont redimensionnées pour être réduites sans être redessinées. |
|Use a white fill for accessibility. Most objects in your icons will require a white background to be legible across Office UI themes and in high-contrast modes.  |Évitez de vous fier à votre logo ou marque pour communiquer ce que fait une commande de complément. Les repères de marque ne sont pas toujours reconnaissables sur des icônes de petites tailles et lorsque des modificateurs sont appliqués. Les marques de marque sont souvent en conflit avec les styles d’icône du ruban de l’application Office et peuvent faire concurrence à l’attention des utilisateurs dans un environnement saturé. |
|Utilisez le format PNG avec un arrière-plan transparent. |*Aucun.*|
|Évitez le contenu localisable dans les icônes, y compris les caractères typographiques, les paragraphes en drapeau et les points d’interrogation. |*Aucun.*|

## <a name="icon-size-recommendations-and-requirements"></a>Configuration requise et recommandations sur la taille des icônes

Les icônes du bureau Office sont des images bitmap. Différentes tailles apparaissent en fonction du paramètre PPP de l’utilisateur et du mode tactile. Incluez les huit tailles prises en charge pour créer la meilleure expérience possible dans tous les contextes et résolutions pris en charge. Les tailles prises en charge sont les suivantes : trois sont requises.

- 16 px (obligatoire)
- 20 px
- 24 px
- 32 px (obligatoire)
- 40 px
- 48 px
- 64 px (recommandé, meilleur choix pour Mac)
- 80 px (obligatoire)

> [!IMPORTANT]
> Pour obtenir une image représentant l’icône représentative de votre complément, consultez [Créer des descriptions efficaces dans AppSource et dans Office](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) pour connaître la taille et d’autres exigences.

Veillez à renouveler les icônes pour chaque taille au lieu de les réduire pour les ajuster.

![Illustration de la recommandation de redessiner les icônes par taille plutôt que de réduire les icônes. Par exemple, vous devrez peut-être utiliser moins d’éléments dans une petite icône plutôt que de simplement réduire une image plus grande.](../images/icon-resizing.png)

## <a name="icon-anatomy-and-layout"></a>Mise en page et structure de l’icône

Les icônes Office sont généralement composées d’un élément de base avec des modificateurs d’action et conceptuels superposés. Les modificateurs d’action représentent des concepts tels qu’ajouter, ouvrir, nouveau ou fermer. Les modificateurs conceptuels représentent l’état, l’altération ou une description de l’icône.

To create commands that align with the Office UI, follow layout guidelines for the base element and modifiers. This ensures that your commands look professional and that your customers will trust your add-in. If you make exceptions to these guidelines, do so intentionally.

L’image suivante montre la disposition des éléments de base et modificateurs dans une icône Office.

![Diagramme montrant un élément de base d’icône au centre avec un modificateur en bas à droite et un modificateur d’action en haut à gauche.](../images/icon-layouts.png)

- Éléments de base centraux dans le cadre de pixel avec remplissage vide tout autour.
- Placez les modificateurs d’action dans le coin supérieur gauche.
- Placez les modificateurs conceptuels dans la partie inférieure droite.
- Limitez le nombre d’éléments dans les icônes. À 32 px, limitez le nombre de modificateurs à un maximum de deux. À 16 px, limitez le nombre de modificateurs à un.

### <a name="base-element-padding"></a>Remplissage d’un élément de base

Placez les éléments de base de façon cohérente en fonction des tailles. Si les éléments de base ne peuvent pas être centrés dans le cadre, alignez-les en haut à gauche, en laissant les pixels supplémentaires dans la partie inférieure droite. Pour de meilleurs résultats, appliquez les instructions de remplissage répertoriées dans le tableau de la section suivante.

### <a name="modifiers"></a>Modificateurs

Tous les modificateurs doivent avoir un découpage transparent de 1 px entre chaque élément, y compris l’arrière-plan. Les éléments ne doivent pas se chevaucher directement. Créez des espaces entre les règles et les bords. Les modificateurs peuvent varier légèrement en taille, mais utilisez ces dimensions comme point de départ.

|Taille de l’icône|Remplissage autour de l’élément de base|Taille du modificateur|
|:---|:---|:---|
|16 px|0|9 px|
|20 px|1px|10 px|
|24 px|1px|12 px|
|32 px|2px|14 px|
|40 px|2px|20 px|
|48 px|3px|22 px|
|64 px|5px|29 px|
|80 px|5px|38 px|

## <a name="icon-colors"></a>Couleurs de l’icône

> [!NOTE]
> Les couleurs recommandées concernent les icônes du ruban utilisées dans les [Commandes de complément](add-in-commands.md). Ces icônes ne sont pas affichées avec l’interface utilisateur Fluent et la palette de couleurs est différente de la palette décrite dans [Microsoft UI Fabric | Couleurs | Partagé](https://fluentfabric.azurewebsites.net/#/color/shared).

Les icônes Office ont une palette de couleurs limitée. Utilisez les couleurs répertoriées dans le tableau suivant pour garantir une intégration parfaite avec l’interface utilisateur d’Office. Appliquez les instructions suivantes à l’utilisation de la couleur.

- Use color to communicate meaning rather than for embellishment. It should highlight or emphasize an action, status, or an element that explicitly differentiates the mark.
- Si possible, n’utilisez qu’une seule couleur supplémentaire au-delà du gris. Limitez les couleurs supplémentaires à deux au maximum.
- Les couleurs ont une apparence cohérente dans toutes les tailles d’icône. Les icônes Office ont des palettes de couleurs légèrement différentes pour des tailles d’icônes différentes. Les icônes 16 px et plus petites sont légèrement plus sombres et plus vibrantes que 32 px et plus grandes icônes. Sans ces ajustements discrets, les couleurs semblent varier en taille.

|Nom de la couleur|RVB|Hex|Couleur|Catégorie|
|:---|:---|:---|:---|:---|
|Texte gris (80)|80, 80, 80|#505050| ![Couleur grise 80 pour le texte.](../images/color-text-gray-80.png) |Texte|
|Texte gris (95)|95, 95, 95|#5F5F5F| ![Couleur grise 95 pour le texte.](../images/color-text-gray-95.png) |Texte|
|Texte gris (105)|105, 105, 105|#696969| ![Couleur grise 105 pour le texte.](../images/color-text-gray-105.png) |Texte|
|Gris foncé 32|128, 128, 128|#808080| ![Couleur gris foncé pour 32 px et plus.](../images/color-dark-gray-32.png) |32 px et versions ultérieures|
|Gris moyen 32|158, 158, 158|#9E9E9E| ![Couleur grise moyenne pour 32 px et plus.](../images/color-medium-gray-32.png) |32 px et versions ultérieures|
|TOUT gris clair|179, 179, 179|#B3B3B3| ![Couleur gris clair pour toutes les tailles d’image.](../images/color-light-gray-all.png) |Toutes les tailles|
|Gris foncé 16|114, 114, 114|#727272| ![Couleur gris foncé pour 16 px et plus petit.](../images/color-dark-gray-16.png) |16 px et versions inférieures|
|Gris moyen 16|144, 144, 144|#909090| ![Couleur grise moyenne pour 16 px et plus petit.](../images/color-medium-gray-16.png) |16 et moins|
|Bleu 32|77, 130, 184|#4d82B8| ![Couleur bleue pour 32 px et plus.](../images/color-blue-32.png) |32 px et versions ultérieures|
|Bleu 16|74, 125, 177|#4A7DB1| ![Couleur bleue pour 16 px et plus petit.](../images/color-blue-16.png) |16 px et versions inférieures|
|TOUT jaune|234, 194, 130|#EAC282| ![Couleur jaune pour toutes les tailles d’image.](../images/color-yellow-all.png) |Toutes les tailles|
|Orange 32|231, 142, 70|#E78E46| ![Couleur orange pour 32 px et plus.](../images/color-orange-32.png) |32 px et versions ultérieures|
|Orange 16|227, 142, 70|#E3751C| ![Couleur orange pour 16 px et plus petit.](../images/color-orange-16.png) |16 px et versions inférieures|
|TOUT rose|230, 132, 151|#E68497| ![Couleur rose pour toutes les tailles d’image.](../images/color-pink-all.png) |Toutes les tailles|
|Vert 32|118, 167, 151|#76A797| ![Couleur verte pour 32 px et plus.](../images/color-green-32.png) |32 px et versions ultérieures|
|Vert 16|104, 164, 144|#68A490| ![Couleur verte pour 16 px et plus petit.](../images/color-green-16.png) |16 px et versions inférieures|
|Rouge 32|216, 99, 68|#D86344| ![Couleur rouge pour 32 px et plus.](../images/color-red-32.png) |32 px et versions ultérieures|
|Rouge 16|214, 85, 50|#D65532| ![Couleur rouge pour 16 px et plus petit.](../images/color-red-16.png) |16 px et versions inférieures|
|Violet 32|152, 104, 185|#9868B9| ![Couleur pourpre pour 32 px et plus.](../images/color-purple-32.png) |32 px et versions ultérieures|
|Violet 16|137, 89, 171|#8959AB| ![Couleur violet pour 16 px et plus petit.](../images/color-purple-16.png) |16 px et versions inférieures|

## <a name="icons-in-high-contrast-modes"></a>Icônes en modes de contraste élevé

Les icônes Office sont conçues pour un rendu correct en mode de contraste élevé. Les éléments de premier plan sont bien différenciés des arrière-plans pour optimiser la lisibilité et permettre le recoloriage. En modes de contraste élevé, Office recolorie tous les pixels de votre icône avec une valeur rouge, verte ou bleue inférieure à 190 en noir plein. Tous les autres pixels sont blancs. Autrement dit, chaque canal RVB est évalué lorsque les valeurs 0-189 sont noires et les valeurs 190-255 sont blanches. D’autres thèmes à contraste élevé recolorient à l’aide du même seuil de valeur 190 mais avec des règles différentes. Par exemple, le thème blanc à contraste élevé recolorie tous les pixels supérieurs à 190 en opaque, mais tous les autres pixels en transparent. Appliquez les instructions suivantes pour optimiser la lisibilité dans les paramètres à contraste élevé.

- Essayez de différencier les éléments de premier plan et d’arrière-plan par rapport au seuil de valeur 190.
- Suivez les styles visuels des icônes Office.
- Utilisez des couleurs de notre palette d’icônes.
- Évitez d’utiliser des dégradés.
- Évitez les grands blocs de couleur avec des valeurs similaires.

## <a name="see-also"></a>Voir aussi

- [Élément de manifeste d’icône](/javascript/api/manifest/icon)
- [Élément de manifeste IconUrl](/javascript/api/manifest/iconurl)
- [Élément manifeste HighResolutionIconUrl](/javascript/api/manifest/highresolutioniconurl)
- [Créer une icône pour votre complément](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in)
