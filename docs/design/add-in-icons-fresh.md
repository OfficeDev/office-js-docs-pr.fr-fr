---
title: Instructions relatives aux icônes de style frais pour les compléments Office
description: Obtenez des instructions sur l’utilisation des icônes d’icône de style fraîche dans les compléments Office.
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: d6acd2b0b17be7b00f14d63c73714c6dc83d45b7
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132206"
---
# <a name="fresh-style-icon-guidelines-for-office-add-ins"></a>Instructions relatives aux icônes de style frais pour les compléments Office

Les versions Office 2013 + (sans abonnement) d’Office utilisent les nouveaux styles de style iconographie. Si vous préférez que vos icônes correspondent au style monoligne de Microsoft 365, reportez-vous à la rubrique [règles d’icône de style monoligne pour les compléments Office](add-in-icons-monoline.md).

## <a name="office-fresh-visual-style"></a>Nouveau style visuel Office

Les icônes nouvelles incluent uniquement les éléments de communication essentiels. Les éléments non essentiels, tels que la source de lumière, les dégradés et les perspectives, sont supprimés. Les icônes simplifiées prennent en charge l’analyse rapide des commandes et des contrôles. Suivez ce style pour ajuster au mieux les clients sans abonnement Office.

## <a name="best-practices"></a>Meilleures pratiques

Suivez ces instructions lorsque vous créez vos icônes :

|À faire|À ne pas faire|
|:---|:---|
|Restez simple et clair en vous concentrant sur les éléments clés de la communication.| N’utilisez pas d’artefacts qui rendent votre icône désordonnée.|
|Utilisez le langage d’icône Office pour représenter des comportements ou des concepts.|Ne réactivez pas les glyphes de structure d’interface utilisateur Office pour les commandes de complément dans le ruban d’application Office ou les menus contextuels. Les icônes de structure sont stylistiquement différentes et ne correspondront pas.|
|Réutilisez les métaphores visuelles d’Office courantes telles que le pinceau pour mettre en forme ou la loupe pour rechercher.|Ne réutilisez pas les métaphores visuelles pour différentes commandes. L’utilisation de la même icône pour différents comportements et concepts peut semer la confusion. |
|Redessinez vos icônes pour les réduire ou les agrandir. Prenez le temps de redessiner les découpages, les coins et des bords arrondis pour optimiser la netteté de ligne. |Ne redimensionnez pas vos icônes en réduisant ou en agrandissant leurs tailles. Cela peut entraîner une mauvaise qualité visuelle et des actions peu claires. Les icônes complexes créées dans une plus grande taille risquent de perdre en clarté si elles sont redimensionnées pour être réduites sans être redessinées. |
|Utilisez un remplissage blanc pour améliorer l’accessibilité. La plupart des objets dans les icônes nécessitent un arrière-plan blanc pour être lisibles sur les thèmes de l’interface utilisateur d’Office et en mode contraste élevé.  |Évitez de vous fier à votre logo ou marque pour communiquer ce que fait une commande de complément. Les repères de marque ne sont pas toujours reconnaissables sur des icônes de petites tailles et lorsque des modificateurs sont appliqués. Les marques de marque sont souvent en conflit avec les styles d’icône du ruban de l’application Office et peuvent rivaliser pour attirer l’attention des utilisateurs dans un environnement saturé. |
|Utilisez le format PNG avec un arrière-plan transparent. ||
|Évitez le contenu localisable dans les icônes, y compris les caractères typographiques, les paragraphes en drapeau et les points d’interrogation. ||

## <a name="icon-size-recommendations-and-requirements"></a>Configuration requise et recommandations sur la taille des icônes

Les icônes du bureau Office sont des images bitmap. Différentes tailles apparaissent en fonction du paramètre PPP de l’utilisateur et du mode tactile. Incluez les huit tailles prises en charge pour créer la meilleure expérience possible dans tous les contextes et résolutions pris en charge. Voici les tailles prises en charge - trois sont obligatoires :

- 16 px (obligatoire)
- 20 px
- 24 px
- 32 px (obligatoire)
- 40 px
- 48 px
- 64 px (recommandé, meilleur choix pour Mac)
- 80 px (obligatoire)

Veillez à renouveler les icônes pour chaque taille au lieu de les réduire pour les ajuster.

![Illustration de la recommandation pour redessiner des icônes par taille au lieu de réduire les icônes. Par exemple, vous pouvez avoir besoin d’utiliser moins d’éléments dans une petite icône plutôt que de mettre à l’échelle une image plus grande.](../images/icon-resizing.png)

<!--
The following table shows the icon sizes that render for different modes at different DPI settings.

|DPI |**Small**||**Medium**||**Large**||**Extra large**|
|:---|:---|:---|:---|:---|:---|:---|:---|
|    |**Mouse**|**Touch**|**Mouse**|**Touch**|**Mouse**|**Touch**|-|
|100%|16px|20px|24px||32px|40px|48px|
|125%|20px|24px|||40px|48px|60px|
|150%|24px|24px|36px||48px|48px|72px|
|200%|32px|40px|48px||64px|80px|96px|
|250%|40px||||80px||120px|
|300%|48px||||96px||144px

> [!NOTE]
> At DPI settings of 150% or greater, the icon does not get swapped out for a larger size when Touch mode is engaged. At DPI settings greater than 250%, Touch mode is turned off by default.

The following table lists the locations for certain icon sizes.

|Location|100% DPI|200% DPI|250% DPI|
|:-------|:-------|:-------|:-------|
|Small ribbon button|16px|32px|40px|
|Contextual menu|16px|32px|40px|
|Quick access toolbar (QAT)|16px|32px|40px|
|Large ribbon icon|32px|64px|80px|

-->

## <a name="icon-anatomy-and-layout"></a>Mise en page et structure de l’icône

Les icônes Office sont généralement composées d’un élément base avec des modificateurs d’action et conceptuels superposés. Les modificateurs d’action représentent des concepts tels qu’ajouter, ouvrir, nouveau ou fermer. Les modificateurs conceptuels représentent l’état, l’altération ou une description de l’icône.

Pour créer des commandes qui s’alignent sur l’interface utilisateur d’Office, suivez les instructions de mise en forme pour les éléments de base et les modificateurs. Cela garantit que vos commandes auront un aspect professionnel et que vos clients auront confiance en votre complément. Si vous apportez des exceptions à ces instructions, faites-le intentionnellement.

L’image suivante montre la disposition des éléments de base et modificateurs dans une icône Office.

![Diagramme montrant un élément de base d’icône dans le centre avec un modificateur sur l’angle inférieur droit et un modificateur d’action dans le coin supérieur gauche](../images/icon-layouts.png)

- Éléments de base centraux dans le cadre de pixel avec remplissage vide tout autour.
- Placez les modificateurs d’action dans le coin supérieur gauche.
- Placez les modificateurs conceptuels dans la partie inférieure droite.
- Limitez le nombre d’éléments dans les icônes. À 32 PX, limitez le nombre de modificateurs à un maximum de deux. À 16 PX, limitez le nombre de modificateurs à un.

### <a name="base-element-padding"></a>Remplissage d’un élément de base

Placez les éléments de base de façon cohérente en fonction des tailles. Si les éléments de base ne peuvent pas être centrés dans le cadre, alignez-les en haut à gauche, en laissant les pixels supplémentaires dans la partie inférieure droite. Pour obtenir de meilleurs résultats, appliquez les instructions de remplissage indiquées dans le tableau de la section suivante.

### <a name="modifiers"></a>Modificateurs

Tous les modificateurs doivent disposer d’un découpage transparent de 1 px entre chaque élément, y compris l’arrière-plan. Les éléments ne doivent pas se chevaucher directement. Créez des espaces entre les règles et les bords. Les modificateurs peuvent varier légèrement en taille, mais utilisez ces dimensions comme point de départ.

|Taille de l’icône|Remplissage autour de l’élément de base|Taille du modificateur|
|:---|:---|:---|
|16 PX|0|9 PX|
|20 px|1px|10 px|
|24 px|1px|12 px|
|32 PX|2px|14 px|
|40 px|2px|20 px|
|48 px|3px|22 PX|
|64 px|5px|29 PX|
|80 PX|5px|38 PX|

## <a name="icon-colors"></a>Couleurs de l’icône

> [!NOTE]
> Les couleurs recommandées concernent les icônes du ruban utilisées dans les [Commandes de complément](add-in-commands.md). Ces icônes ne sont pas restituées dans Microsoft UI Fabric et la palette de couleurs est différente de la palette présentée dans [Microsoft UI Fabric | Couleurs | Partagé](https://fluentfabric.azurewebsites.net/#/color/shared).

Les icônes Office ont une palette de couleurs limitée. Utilisez les couleurs répertoriées dans le tableau suivant pour garantir une intégration parfaite avec l’interface utilisateur d’Office. Appliquez les instructions suivantes sur l’utilisation des couleurs :

- Utilisez la couleur pour véhiculer une signification plutôt que pour embellir. Elle doit mettre en surbrillance ou mettre en évidence une action, un état ou un élément qui différencie explicitement le repère.
- Si possible, n’utilisez qu’une seule couleur supplémentaire au-delà du gris. Limitez les couleurs supplémentaires à deux au maximum.
- Les couleurs ont une apparence cohérente dans toutes les tailles d’icône. Les icônes Office ont des palettes de couleurs légèrement différentes pour des tailles d’icônes différentes. les icônes 16 PX et plus petites sont légèrement plus foncées et plus éclatantes que 32 PX et des icônes plus grandes. Sans ces ajustements discrets, les couleurs semblent varier en taille.

|Nom de la couleur|RVB|Hex|Couleur|Catégorie|
|:---|:---|:---|:---|:---|
|Texte gris (80)|80, 80, 80|#505050| ![Gris 80 couleur pour le texte](../images/color-text-gray-80.png) |Texte|
|Texte gris (95)|95, 95, 95|#5F5F5F| ![Gris 95 couleur pour le texte](../images/color-text-gray-95.png) |Texte|
|Texte gris (105)|105, 105, 105|#696969| ![Gris 105 couleur pour le texte](../images/color-text-gray-105.png) |Texte|
|Gris foncé 32|128, 128, 128|#808080| ![Couleur gris foncé pour 32 PX et plus grand](../images/color-dark-gray-32.png) |32 PX et versions supérieures|
|Gris moyen 32|158, 158, 158|#9E9E9E| ![Couleur gris moyen de 32 PX et supérieure](../images/color-medium-gray-32.png) |32 PX et versions supérieures|
|TOUT gris clair|179, 179, 179|#B3B3B3| ![Couleur gris clair pour toutes les tailles d’image](../images/color-light-gray-all.png) |Toutes les tailles|
|Gris foncé 16|114, 114, 114|#727272| ![Couleur gris foncé pour 16 PX et plus petite](../images/color-dark-gray-16.png) |16 PX et inférieur|
|Gris moyen 16|144, 144, 144|#909090| ![Couleur gris moyen pour 16 PX et plus petite](../images/color-medium-gray-16.png) |16 et moins|
|Bleu 32|77, 130, 184|#4d82B8| ![Couleur bleue pour 32 PX et plus grande](../images/color-blue-32.png) |32 PX et versions supérieures|
|Bleu 16|74, 125, 177|#4A7DB1| ![Couleur bleue pour 16 PX et plus petite](../images/color-blue-16.png) |16 PX et inférieur|
|TOUT jaune|234, 194, 130|#EAC282| ![Couleur jaune pour toutes les tailles d’image](../images/color-yellow-all.png) |Toutes les tailles|
|Orange 32|231, 142, 70|#E78E46| ![Couleur orange pour 32 PX et plus grande](../images/color-orange-32.png) |32 PX et versions supérieures|
|Orange 16|227, 142, 70|#E3751C| ![Couleur orange pour 16 PX et plus petite](../images/color-orange-16.png) |16 PX et inférieur|
|TOUT rose|230, 132, 151|#E68497| ![Couleur rose pour toutes les tailles d’image](../images/color-pink-all.png) |Toutes les tailles|
|Vert 32|118, 167, 151|#76A797| ![Couleur verte pour 32 PX et plus grande](../images/color-green-32.png) |32 PX et versions supérieures|
|Vert 16|104, 164, 144|#68A490| ![Couleur verte pour 16 PX et plus petite](../images/color-green-16.png) |16 PX et inférieur|
|Rouge 32|216, 99, 68|#D86344| ![Couleur rouge pour 32 PX et plus grande](../images/color-red-32.png) |32 PX et versions supérieures|
|Rouge 16|214, 85, 50|#D65532| ![Couleur rouge pour 16 PX et plus petite](../images/color-red-16.png) |16 PX et inférieur|
|Violet 32|152, 104, 185|#9868B9| ![Couleur violette pour 32 PX et plus grande](../images/color-purple-32.png) |32 PX et versions supérieures|
|Violet 16|137, 89, 171|#8959AB| ![Couleur violette pour 16 PX et plus petite](../images/color-purple-16.png) |16 PX et inférieur|

## <a name="icons-in-high-contrast-modes"></a>Icônes en modes de contraste élevé

Les icônes Office sont conçues pour un rendu correct en mode de contraste élevé. Les éléments de premier plan sont bien différenciés des arrière-plans pour optimiser la lisibilité et permettre le recoloriage. En modes de contraste élevé, Office recolorie tous les pixels de votre icône avec une valeur rouge, verte ou bleue inférieure à 190 en noir plein. Tous les autres pixels sont blancs. Autrement dit, chaque canal RVB est évalué lorsque les valeurs 0-189 sont noires et les valeurs 190-255 sont blanches. D’autres thèmes à contraste élevé recolorient à l’aide du même seuil de valeur 190 mais avec des règles différentes. Par exemple, le thème blanc à contraste élevé recolorie tous les pixels supérieurs à 190 en opaque, mais tous les autres pixels en transparent. Appliquez les instructions suivantes pour optimiser la lisibilité dans les paramètres de contraste élevé :

- Essayez de différencier les éléments de premier plan et d’arrière-plan par rapport au seuil de valeur 190.
- Suivez les styles visuels des icônes Office.
- Utilisez des couleurs de notre palette d’icônes.
- Évitez d’utiliser des dégradés.
- Évitez les grands blocs de couleur avec des valeurs similaires.
