# <a name="icons"></a>Icônes
Les icônes sont la représentation visuelle d'un comportement ou d'un concept. Elles sont souvent utilisées pour ajouter du sens aux contrôles et aux commandes. Les objets visuels, qu'ils soient réalistes ou symboliques, permettent à l'utilisateur de naviguer à travers l'interface utilisateur de la même manière que les signes aident les utilisateurs à naviguer dans leur environnement. Ils doivent être simples, clairs et ne contenir que les détails nécessaires pour permettre aux clients d'analyser rapidement l'action qui se produira lorsqu'ils choisiront un contrôle.

Les interfaces de ruban Office ont un style visuel standard. Cela garantit la cohérence et la familiarité de toutes les applications Office. Les instructions vous aideront à concevoir un ensemble d'actifs PNG pour votre solution qui s'intègre naturellement à Office.

De nombreux conteneurs HTML contiennent des contrôles avec iconographie. Utilisez la police personnalisée d’Office UI Fabric pour le rendu des icônes de style Office dans votre complément. La police d’icône de Fabric contient de nombreux glyphes pour les métaphores Office courantes que vous pouvez redimensionner, colorier et personnaliser selon vos besoins. Si vous avez un langage visuel existant avec votre propre jeu d’icônes, n’hésitez pas à l’utiliser dans vos canevas HTML. Créer la continuité avec votre marque avec un jeu d’icônes standard est une partie importante de tout langage de création. Soyez prudent pour éviter de créer de la confusion pour les clients en conflit avec les métaphores Office.


# <a name="design-icons-for-add-in-commands"></a>Concevoir des icônes pour les commandes de complément

[Commandes de complément](add-in-commands.md) Ajoutez des boutons, du texte et des icônes à l’interface utilisateur Office. Vos boutons de commande de complément doivent fournir des icônes significatives et des étiquettes qui identifient clairement l’action que l’utilisateur effectue lorsqu’il utilise une commande. Cet article fournit des instructions stylistiques et de production pour vous aider à concevoir des icônes s’intégrant parfaitement avec Office. 

## <a name="office-icon-design-principles"></a>Principes de conception des icônes Office

La version Office 2013 des clients de bureau Office inclut une iconographie actualisée. La modification stylistique de remplacement est une réduction. Les nouvelles icônes incluent uniquement les éléments de communication essentiels. Les éléments non essentiels, tels que la source de lumière, les dégradés et les perspectives, sont supprimés. Les icônes simplifiées prennent en charge l’analyse rapide des commandes et des contrôles. Suivez ce style pour mieux correspondre à Office.

Les icônes Office sont basées sur les principes de conception suivants : 

- Interprétation moderne de la collection d’icônes Office 
- À la fois nouveau et familier  
- Simple, clair et direct 

L’image suivante montre les icônes qui appliquent les principes de conception modernes.

![Image illustrant les anciennes icônes Office et l’interprétation moderne actualisée des icônes](../images/icons-images.png)

## <a name="best-practices"></a>Meilleures pratiques

Suivez ces instructions lorsque vous créez vos icônes : 

|À faire|À ne pas faire|
|:---|:---|
|Gardez les éléments visuels simples et clairs, en mettant l'accent sur l'élément clé de la communication.| N'utilisez pas d'artefacts qui rendent votre icône désordonnée.|
|Utilisez le langage d’icône Office pour représenter des comportements ou des concepts.|Ne redéfinissez pas les glyphes Office UI Fabric pour les commandes de complément dans le ruban Office ou les menus contextuels. Les icônes Fabric sont stylistiquement différentes et ne correspondront pas.|
|Réutilisez les métaphores visuelles d’Office courantes telles que le pinceau pour mettre en forme ou la loupe pour rechercher.|Ne réutilisez pas les métaphores visuelles pour différentes commandes. Utiliser la même icône pour différents comportements et concepts peut prêter à confusion. |
|Redessinez vos icônes pour les rendre petites ou plus grandes. Prenez le temps de redessiner les découpes, les coins et les bords arrondis pour agrandir la clarté de la ligne. |Ne redimensionnez pas vos icônes en les rétrécissant ou en les agrandissant. Cela peut entraîner une mauvaise qualité visuelle et des actions imprécises. Les icônes complexes créées à une plus grande taille peuvent perdre de leur clarté si elles sont redimensionnées pour être plus petites sans redessiner. |
|Utilisez un remplissage blanc pour améliorer l’accessibilité. La plupart des objets dans les icônes nécessitent un arrière-plan blanc pour être lisibles sur les thèmes de l’interface utilisateur d’Office et en mode contraste élevé.  ||
|Utilisez le format PNG avec un arrière-plan transparent. ||
|Évitez du contenu de localisation dans les icônes, y compris les caractères typographiques, les indications de paragraphes en drapeau et les points d’interrogation. ||



## <a name="icon-size-recommendations-and-requirements"></a>Configuration requise et recommandations sur la taille des icônes

Les icônes du bureau Office 2016 sont des images bitmap. Différentes tailles apparaissent en fonction du paramètre PPP de l’utilisateur et du mode tactile. Incluez les huit tailles prises en charge pour créer la meilleure expérience possible dans tous les contextes et résolutions pris en charge. Voici les tailles prises en charge - trois sont obligatoires :

- 16 px (obligatoire)
- 20 px
- 24 px
- 32 px (obligatoire)
- 40 px
- 48 px
- 64 px (recommandé, meilleur choix pour Mac)
- 80 px (obligatoire)  

Veillez à renouveler les icônes pour chaque taille au lieu de les réduire pour les ajuster.

![Illustration présentant la recommandation qui indique de redimensionner les icônes plutôt que de les réduire](../images/icon-resizing.png)

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

Les icônes Office sont généralement constituées d’un élément de base avec des modificateurs d’action et conceptuels superposés. Les modificateurs d’action représentent des concepts tels qu’ajouter, ouvrir, nouveau ou fermer. Les modificateurs conceptuels représentent l’état, l’altération ou une description de l’icône. 

Pour créer des commandes qui s’alignent sur l’interface utilisateur d’Office, suivez les instructions de mise en forme pour les éléments de base et les modificateurs. Cela garantit que vos commandes auront un aspect professionnel et que vos clients auront confiance en votre complément. Si vous apportez des exceptions à ces instructions, faites-le intentionnellement.

L’image suivante montre la disposition des éléments de base et modificateurs dans une icône Office.

![Image illustrant un élément de base d’icône dans le centre avec un modificateur dans le coin inférieur droit et un modificateur d’action dans le coin supérieur gauche](../images/icon-layouts.png)

- Éléments de base centraux dans le cadre de pixel avec remplissage vide tout autour.
- Placez les modificateurs d’action dans le coin supérieur gauche. 
- Placez les modificateurs conceptuels dans la partie inférieure droite.
- Limitez le nombre d’éléments dans les icônes. En 32 px, limitez le nombre de modificateurs à un maximum de deux. En 16 px, limitez le nombre de modificateurs à un.

###<a name="base-element-padding"></a>Marge intérieure de l'élément de base
Placez les éléments de base de façon cohérente en fonction des tailles. Si les éléments de base ne peuvent pas être centrés dans le cadre, alignez-les en haut à gauche, en laissant les pixels supplémentaires dans la partie inférieure droite. Pour obtenir de meilleurs résultats, appliquez les instructions de remplissage répertoriées dans le tableau suivant.

###<a name="modifiers"></a>Modificateurs
Tous les modificateurs doivent avoir un découpage transparent 1 px entre chaque élément, y compris l’arrière-plan. Les éléments ne doivent pas se chevaucher directement. Créez des espaces entre les règles et les bords. Les modificateurs peuvent varier légèrement en taille, mais utilisez ces dimensions comme point de départ.


|**Taille de l’icône**|**Remplissage autour de l’élément de base**|**Taille du modificateur**|
|:---|:---|:---|
|16 px|0|9 px|
|20 px|1 px|10 px|
|24 px|1 px|12 px|
|32 px|2 px|14 px|
|40 px|2 px|20 px|
|48 px|3 px|22 px|
|64 px|5 px|29 px|
|80 px|5 px|38 px|


## <a name="icon-colors"></a>Couleurs de l’icône

Les icônes Office ont une palette de couleurs limitée. Utilisez les couleurs répertoriées dans le tableau suivant pour garantir une intégration parfaite avec l’interface utilisateur d’Office. Appliquez les instructions suivantes sur l’utilisation des couleurs : 

- Utilisez la couleur pour véhiculer une signification plutôt que pour embellir. Elle doit mettre en surbrillance ou mettre en évidence une action, un état ou un élément qui différencie explicitement le repère.  
- Si possible, n’utilisez qu’une seule couleur supplémentaire au-delà du gris. Limitez les couleurs supplémentaires à deux au maximum.
- Les couleurs ont une apparence cohérente dans toutes les tailles d’icône. Les icônes Office ont des palettes de couleurs légèrement différentes pour des tailles d’icônes différentes. Les icônes 16 px et plus petites sont légèrement plus sombres et plus percutantes que les icônes 32 px et plus grandes. Sans ces ajustements discrets, les couleurs semblent varier en taille.   

|**Nom de la couleur**|**RVB**|**Hexadécimal**|**Couleur**|**Catégorie**|
|:---|:---|:---|:---|:---|
|Texte gris (80)|80, 80, 80|#505050| ![Image couleur texte gris 80](../images/color-text-gray-80.png) |Texte|
|Texte gris (95)|95, 95, 95|#5F5F5F| ![Image couleur texte gris 95](../images/color-text-gray-95.png) |Texte|
|Texte gris (105)|105, 105, 105|#696969| ![Image couleur texte gris 105](../images/color-text-gray-105.png) |Texte|
|Gris foncé 32|128, 128, 128|#808080| ![Image couleur gris foncé 32](../images/color-dark-gray-32.png) |32 et plus|
|Gris moyen 32|158, 158, 158|#9E9E9E| ![Image couleur gris moyen 32](../images/color-medium-gray-32.png) |32 et plus|
|TOUT gris clair|179, 179, 179|#B3B3B3| ![Image couleur tout en gris clair](../images/color-light-gray-all.png) |Toutes les tailles|
|Gris foncé 16|114, 114, 114|#727272| ![Image couleur gris foncé 16](../images/color-dark-gray-16.png) |16 et moins|
|Gris moyen 16|144, 144, 144|#909090| ![Image couleur gris moyen 16](../images/color-medium-gray-16.png) |16 et moins|
|Bleu 32|77, 130, 184|#4d82B8| ![Image couleur bleu 32](../images/color-blue-32.png) |32 et plus|
|Bleu 16|74, 125, 177|#4A7DB1| ![Image couleur bleu 16](../images/color-blue-16.png) |16 et moins|
|TOUT jaune|234, 194, 130|#EAC282| ![Image couleur tout en jaune](../images/color-yellow-all.png) |Toutes les tailles|
|Orange 32|231, 142, 70|#E78E46| ![Image couleur orange 32](../images/color-orange-32.png) |32 et plus|
|Orange 16|227, 142, 70|#E3751C| ![Image couleur orange 16](../images/color-orange-16.png) |16 et moins|
|TOUT rose|230, 132, 151|#E68497| ![Image couleur tout en rose](../images/color-pink-all.png) |Toutes les tailles|
|Vert 32|118, 167, 151|#76A797| ![Image couleur vert 32](../images/color-green-32.png) |32 et plus|
|Vert 16|104, 164, 144|#68A490| ![Image couleur vert 16](../images/color-green-16.png) |16 et moins|
|Rouge 32|216, 99, 68|#D86344| ![Image couleur rouge 32](../images/color-red-32.png) |32 et plus|
|Rouge 16|214, 85, 50|#D65532| ![Image couleur rouge 16](../images/color-red-16.png) |16 et moins|
|Violet 32|152, 104, 185|#9868B9| ![Image couleur violet 32](../images/color-purple-32.png) |32 et plus|
|Violet 16|137, 89, 171|#8959AB| ![Image couleur violet 16](../images/color-purple-16.png) |16 et moins|


## <a name="icons-in-high-contrast-modes"></a>Icônes en modes de contraste élevé

Les icônes Office sont conçues pour un rendu correct en mode de contraste élevé. Les éléments de premier plan sont bien différenciés des arrière-plans pour optimiser la lisibilité et permettre le recoloriage. En modes de contraste élevé, Office recolorie tous les pixels de votre icône avec une valeur rouge, verte ou bleue inférieure à 190 en noir plein. Tous les autres pixels sont blancs. Autrement dit, chaque canal RVB est évalué lorsque les valeurs 0-189 sont noires et les valeurs 190-255 sont blanches. D’autres thèmes à contraste élevé recolorient à l’aide du même seuil de valeur 190 mais avec des règles différentes. Par exemple, le thème blanc à contraste élevé recolorie tous les pixels supérieurs à 190 en opaque, mais tous les autres pixels en transparent. Appliquez les instructions suivantes pour optimiser la lisibilité dans les paramètres de contraste élevé :

- Essayez de différencier les éléments de premier plan et d’arrière-plan par rapport au seuil de valeur 190.
- Suivez les styles visuels des icônes Office.
- Utilisez des couleurs de notre palette d’icônes.
- Évitez d’utiliser des dégradés.
- Évitez les grands blocs de couleur avec des valeurs similaires.

## <a name="see-also"></a>Voir aussi

- [Bonnes pratiques en matière de développement de compléments](../concepts/add-in-development-best-practices.md)
- [Commandes de complément pour Excel, Word et PowerPoint](../design/add-in-commands.md)




- Évitez de vous fier à votre logo ou marque pour communiquer ce que fait une commande de complément. Les repères de marque ne sont pas toujours reconnaissables sur des icônes de petites tailles et lorsque des modificateurs sont appliqués. Les repères de marque entrent souvent en conflit avec les styles d’icônes du ruban Office et peuvent gêner l’attention de l’utilisateur dans un environnement saturé.


