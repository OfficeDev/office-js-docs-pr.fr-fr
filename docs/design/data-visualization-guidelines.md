---
title: Règles de style de visualisation de données pour les compléments Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 27de6b6b2f4352488ad8f63c3b6e1250cbfbb324
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945791"
---
# <a name="data-visualization-style-guidelines-for-office-add-ins"></a><span data-ttu-id="a8897-102">Règles de style de visualisation de données pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="a8897-102">Data visualization style guidelines for Office Add-ins</span></span>

<span data-ttu-id="a8897-p101">Une bonne visualisation des données permet aux utilisateurs de rechercher des informations dans leurs données. Ils peuvent utiliser ces informations pour raconter des histoires qui informent et persuadent. Cet article fournit des instructions pour vous aider à créer des visualisations de données efficaces dans vos compléments pour Excel et d’autres applications Office.</span><span class="sxs-lookup"><span data-stu-id="a8897-p101">Good data visualizations help users find insights in their data. They can use those insights to tell stories that inform and persuade. This article provides guidelines to help you design effective data visualizations in your add-ins for Excel and other Office apps.</span></span>

<span data-ttu-id="a8897-p102">Nous vous recommandons d’utiliser [Office UI Fabric](https://developer.microsoft.com/fabric) pour créer l’apparence de vos visualisations de données. Office UI Fabric comprend des styles et des composants qui s’intègrent parfaitement à l’apparence Office.</span><span class="sxs-lookup"><span data-stu-id="a8897-p102">We recommend that you use [Office UI Fabric](https://developer.microsoft.com/fabric) to create the chrome for your data visualizations. Office UI Fabric includes styles and components that integrate seamlessly with the Office look and feel.</span></span> 

<!--The following figure shows a data visualization in an add-in that uses Fabric.

![Image of a data visualization with Fabric elements applied**](../images/fabric-data-visualization.png) 

-->

## <a name="data-visualization-elements"></a><span data-ttu-id="a8897-108">Éléments de visualisation de données</span><span class="sxs-lookup"><span data-stu-id="a8897-108">Data visualization elements</span></span>

<span data-ttu-id="a8897-109">Les visualisations de données partagent un cadre général et des éléments visuels et interactifs communs, notamment des titres, des étiquettes et des tracés de données, comme indiqué dans les figures suivantes.</span><span class="sxs-lookup"><span data-stu-id="a8897-109">Data visualizations share a general framework and common visual and interactive elements, including titles, labels, and data plots, as shown in the following figures.</span></span>

<span data-ttu-id="a8897-110">![Image d’un graphique en courbes avec titre, axes, légende et zone de traçage avec étiquettes](../images/data-visualization-line-chart.png)
![Image d’un histogramme avec axes, quadrillage, légende et tracé de données avec étiquettes](../images/data-visualization-column-chart.png)</span><span class="sxs-lookup"><span data-stu-id="a8897-110">![Image of a line chart with title, axes, legend, and plot area labeled](../images/data-visualization-line-chart.png)
![Image of a column chart with axes, gridlines, legend, and data plot labeled](../images/data-visualization-column-chart.png)</span></span>

### <a name="chart-titles"></a><span data-ttu-id="a8897-111">Titres de graphique</span><span class="sxs-lookup"><span data-stu-id="a8897-111">Chart titles</span></span>

<span data-ttu-id="a8897-112">Suivez ces instructions pour les titres de graphique :</span><span class="sxs-lookup"><span data-stu-id="a8897-112">Follow these guidelines for chart titles:</span></span>

- <span data-ttu-id="a8897-p103">Faites en sorte que vos titres de graphique soient lisibles. Positionnez-les pour créer une hiérarchie visuelle claire par rapport au reste du graphique.</span><span class="sxs-lookup"><span data-stu-id="a8897-p103">Make your chart titles easily readable. Position them to create a clear visual hierarchy in relation to the rest of the chart.</span></span>
- <span data-ttu-id="a8897-p104">En règle générale, utilisez la mise en majuscule de phrase (premier mot en majuscule). Pour créer un contraste ou accentuer des hiérarchies, vous pouvez mettre tout en majuscules, mais faites-le avec parcimonie.</span><span class="sxs-lookup"><span data-stu-id="a8897-p104">In general, use sentence capitalization (capitalize the first word). To create contrast or to reinforce hierarchies, you can use all caps, but all caps should be used sparingly.</span></span>
- <span data-ttu-id="a8897-p105">Intégrez les [recommandations relatives aux polices d’Office UI Fabric](https://developer.microsoft.com/fabric#/styles/typography) pour harmoniser vos graphiques avec l’interface utilisateur Office, qui utilise la police Segoe. Vous pouvez également utiliser une autre police pour différencier le contenu du graphique de l’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a8897-p105">Incorporate the [Office UI Fabric type ramp](https://developer.microsoft.com/fabric#/styles/typography) to make your charts consistent with the Office UI, which uses Segoe. You can also use a different typeface to differentiate chart content from the UI.</span></span>
- <span data-ttu-id="a8897-119">Utilisez des polices sans-serif avec des compteurs de grande taille.</span><span class="sxs-lookup"><span data-stu-id="a8897-119">Use sans-serif typefaces with large counters.</span></span>

<span data-ttu-id="a8897-p106">Les exemples suivants illustrent des polices serif et sans-serif utilisées dans les titres de graphique. Remarquez comment le contraste d’échelle et une utilisation efficace des espaces blancs créent une hiérarchie visuelle forte.</span><span class="sxs-lookup"><span data-stu-id="a8897-p106">The following examples show serif and sans-serif typefaces used in chart titles. Notice how the scale contrast and effective use of white space create a strong visual hierarchy.</span></span>

<span data-ttu-id="a8897-122">![Image d’une visualisation de données avec une police serif](../images/data-visualization-serif.png)
![Image d’une visualisation de données avec une police sans-serif](../images/data-visualization-sans-serif.png)</span><span class="sxs-lookup"><span data-stu-id="a8897-122">![Image of a data visualization with serif font](../images/data-visualization-serif.png)
![Image of a data visualization with sans-serif font](../images/data-visualization-sans-serif.png)</span></span>

### <a name="axis-labels"></a><span data-ttu-id="a8897-123">Étiquettes d’axe</span><span class="sxs-lookup"><span data-stu-id="a8897-123">Axis labels</span></span>

<span data-ttu-id="a8897-p107">Rendez vos étiquettes d’axe suffisamment foncées pour qu’elles soient lisibles, avec des taux de contraste adéquats entre les couleurs de texte et d’arrière-plan. Veillez à ce qu’elles ne soient pas trop foncées pour ne pas se confondre avec l’encre de données.</span><span class="sxs-lookup"><span data-stu-id="a8897-p107">Make your axis labels dark enough to read clearly, with adequate contrast ratios between the text and background colors. Make sure that they are not so dark that they compete with data ink.</span></span>

<span data-ttu-id="a8897-p108">Les nuances de gris clair sont les plus efficaces pour les étiquettes d’axe. Si vous utilisez Fabric, consultez la [palette de couleurs neutres](https://developer.microsoft.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="a8897-p108">Light grays are most effective for axis labels. If you’re using Fabric, see the [Neutral Colors palette](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

### <a name="data-ink"></a><span data-ttu-id="a8897-128">Encre de données</span><span class="sxs-lookup"><span data-stu-id="a8897-128">Data ink</span></span>

<span data-ttu-id="a8897-p109">Les pixels qui représentent les données réelles dans un graphique sont appelés encre de données. Il doit s’agir de l’objectif central de la visualisation. Évitez d’utiliser des ombres portées, des plans lourds ou des éléments de conception inutiles qui faussent ou se confondent avec les données. Utilisez des dégradés uniquement lorsque les valeurs de données sont liées à des valeurs de couleur. Évitez les graphiques en trois dimensions, sauf si une valeur objective mesurable est liée à une troisième dimension.</span><span class="sxs-lookup"><span data-stu-id="a8897-p109">The pixels that represent the actual data in a chart are referred to as data ink. This should be the central focus of the visualization. Avoid the use of drop shadows, heavy outlines, or unnecessary design elements that distort or compete with the data. Use gradients only when data values are tied to color values. Avoid three-dimensional charts unless a measurable, objective value is bound to a third dimension.</span></span>

### <a name="color"></a><span data-ttu-id="a8897-134">Couleur</span><span class="sxs-lookup"><span data-stu-id="a8897-134">Color</span></span>

<span data-ttu-id="a8897-p110">Choisissez des couleurs qui respectent les thèmes du système d’exploitation ou de l’application plutôt que des couleurs codées en dur. En même temps, assurez-vous que les couleurs que vous appliquez ne faussent pas les données. Une utilisation abusive des couleurs dans les visualisations de données peut provoquer une distorsion des données et une lecture incorrecte des informations.</span><span class="sxs-lookup"><span data-stu-id="a8897-p110">Choose colors that follow operating system or application themes rather than hardcoded colors. At the same time, make sure that the colors you apply do not distort the data. Misuse of color in data visualizations can result in data distortion and incorrect reading of information.</span></span>

<span data-ttu-id="a8897-138">Pour des recommandations sur l’utilisation des couleurs dans les visualisations de données, voir les rubriques suivantes :</span><span class="sxs-lookup"><span data-stu-id="a8897-138">For best practices for use of color in data visualizations, see the following:</span></span>


- [<span data-ttu-id="a8897-139">Pourquoi les couleurs de l’arc-en-ciel ne constituent pas la meilleure option pour les visualisations de données ?</span><span class="sxs-lookup"><span data-stu-id="a8897-139">Why rainbow colors aren't the best option for data visualizations</span></span>](http://www.poynter.org/2013/why-rainbow-colors-arent-always-the-best-options-for-data-visualizations/224413/)
- [<span data-ttu-id="a8897-140">Color Brewer 2.0 : Conseils en matière de couleur pour la cartographie</span><span class="sxs-lookup"><span data-stu-id="a8897-140">Color Brewer 2.0: Color Advice for Cartography</span></span>](http://colorbrewer2.org/)
- [<span data-ttu-id="a8897-141">Je veux une teinte</span><span class="sxs-lookup"><span data-stu-id="a8897-141">I Want Hue</span></span>](http://tools.medialab.sciences-po.fr/iwanthue/)

### <a name="gridlines"></a><span data-ttu-id="a8897-142">Quadrillage</span><span class="sxs-lookup"><span data-stu-id="a8897-142">Gridlines</span></span>

<span data-ttu-id="a8897-p111">Le quadrillage est souvent nécessaire pour une lecture précise d’un graphique, mais il doit être présenté comme un élément visuel secondaire, qui améliore l’encre de données, sans se confondre avec elle. Créez un quadrillage statique fin et léger, sauf s’il est conçu spécifiquement pour un contraste élevé. Vous pouvez également utiliser une interaction pour créer un quadrillage dynamique ponctuel qui s’affiche dans le contexte lorsqu’un utilisateur interagit avec un graphique.</span><span class="sxs-lookup"><span data-stu-id="a8897-p111">Gridlines are often necessary for accurately reading a chart, but should be presented as a secondary visual element, enhancing the data ink, not competing with it. Make static gridlines thin and light, unless they are designed specifically for high contrast. You can also use interaction to create dynamic, just-in-time gridlines that appear in context when a user interacts with a chart.</span></span>

<span data-ttu-id="a8897-p112">Les nuances de gris clair sont les plus efficaces pour les quadrillages. Si vous utilisez Fabric, consultez la [palette de couleurs neutres](https://developer.microsoft.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="a8897-p112">Light grays are most effective for gridlines. If you’re using Fabric, see the [Neutral Colors palette](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

<span data-ttu-id="a8897-148">L’image suivante montre une visualisation de données avec le quadrillage.</span><span class="sxs-lookup"><span data-stu-id="a8897-148">The following image shows a data visualization with gridlines.</span></span>

![Image d’une visualisation de données avec le quadrillage](../images/data-visualization-gridlines.png)

### <a name="legends"></a><span data-ttu-id="a8897-150">Légendes</span><span class="sxs-lookup"><span data-stu-id="a8897-150">Legends</span></span>

<span data-ttu-id="a8897-151">Ajoutez des légendes si nécessaire pour :</span><span class="sxs-lookup"><span data-stu-id="a8897-151">Add legends if necessary to:</span></span>

- <span data-ttu-id="a8897-152">Faire une distinction entre les séries</span><span class="sxs-lookup"><span data-stu-id="a8897-152">Distinguish between series</span></span>
- <span data-ttu-id="a8897-153">Présenter des modifications d’échelle ou de valeur</span><span class="sxs-lookup"><span data-stu-id="a8897-153">Present scale or value changes</span></span>

<span data-ttu-id="a8897-p113">Assurez-vous que vos légendes améliorent l’encre de données et ne rivalisent pas avec elle. Placez les légendes :</span><span class="sxs-lookup"><span data-stu-id="a8897-p113">Make sure that your legends enhance the data ink and do not compete with it. Place legends:</span></span>


- <span data-ttu-id="a8897-156">Alignées à gauche, au-dessus de la zone de traçage par défaut, si tous les éléments de légende tiennent au-dessus du graphique.</span><span class="sxs-lookup"><span data-stu-id="a8897-156">Flush left above the plot area by default, if all legend items fit above the chart.</span></span>
- <span data-ttu-id="a8897-157">Dans la partie supérieure droite de la zone de traçage, si tous les éléments de légende ne tiennent pas au-dessus du graphique et ajoutez une zone de texte déroulante, si nécessaire.</span><span class="sxs-lookup"><span data-stu-id="a8897-157">On the upper right side of the plot area, if all legend items do not fit above the chart, and make it scrollable, if necessary.</span></span>

<span data-ttu-id="a8897-p114">Pour optimiser la lisibilité et l’accessibilité, associez des marqueurs de légende à la forme de graphique appropriée. Par exemple, utilisez des marqueurs de légende circulaires pour les légendes de graphique en bulles et de graphique en nuages de points. Utilisez des marques de légende de segment de ligne pour les graphiques en courbes.</span><span class="sxs-lookup"><span data-stu-id="a8897-p114">To optimize for readability and accessibility, map legend markers to the relevant chart shape. For example, use circle legend markers for scatter plot and bubble chart legends. Use line segment legend markers for line charts.</span></span>

### <a name="data-labels-and-tooltips"></a><span data-ttu-id="a8897-161">Info-bulles et étiquettes de données</span><span class="sxs-lookup"><span data-stu-id="a8897-161">Data labels and tooltips</span></span>

<span data-ttu-id="a8897-p115">Assurez-vous que les info-bulles et les étiquettes de données contiennent des variations adéquates d’espace blanc et de type. Utilisez des algorithmes pour réduire l’occlusion et la collision. Par exemple, une info-bulle peut apparaître à droite d’un point de données par défaut, mais à gauche si des bords droits sont détectés.</span><span class="sxs-lookup"><span data-stu-id="a8897-p115">Ensure that data labels and tooltips have adequate white space and type variation. Use algorithms to minimize occlusion and collision. For example, a tooltip might surface to the right of a data point by default, but surface to the left if right edges are detected.</span></span>

## <a name="design-principles"></a><span data-ttu-id="a8897-165">Principes de conception</span><span class="sxs-lookup"><span data-stu-id="a8897-165">Design principles</span></span>

<span data-ttu-id="a8897-166">L’équipe de conception d’Office a élaboré l’ensemble suivant de principes de conception, que nous utilisons lors de la création de visualisations de données pour la suite de produits Office.</span><span class="sxs-lookup"><span data-stu-id="a8897-166">The Office Design team created the following set of design principles, which we use when designing new data visualizations for the Office product suite.</span></span>

### <a name="visual-design-principles"></a><span data-ttu-id="a8897-167">Principes de conception visuelle</span><span class="sxs-lookup"><span data-stu-id="a8897-167">Visual design principles</span></span>

- <span data-ttu-id="a8897-p116">Les visualisations doivent respecter et améliorer les données, facilitant ainsi leur compréhension. Mettez en surbrillance les données, en ajoutant des éléments de soutien uniquement selon les besoins pour fournir un contexte. Évitez les embellissements inutiles (ombres portées, contours, etc.), les éléments de graphique indésirables ou la distorsion des données.</span><span class="sxs-lookup"><span data-stu-id="a8897-p116">Visualizations should honor and enhance the data, making it easy to understand. Highlight the data, adding supporting elements only as needed to provide context. Avoid unnecessary embellishments (drop shadows, outlines, etc), chart junk, or data distortion.</span></span>
- <span data-ttu-id="a8897-p117">Les visualisations doivent encourager l’exploration en fournissant un retour visuel enrichi. Utilisez des modèles d’interaction bien établis, des options d’interface et des commentaires système clairs.</span><span class="sxs-lookup"><span data-stu-id="a8897-p117">Visualizations should encourage exploration by providing rich visual feedback. Use well-established interaction patterns, interface controls, and clear system feedback.</span></span>
- <span data-ttu-id="a8897-p118">Incarnez des principes de conception consacrés. Utilisez les principes de conception de communication typographique et visuelle établis pour améliorer la forme, la lisibilité et le sens.</span><span class="sxs-lookup"><span data-stu-id="a8897-p118">Embody time-honored design principles. Use established typographic and visual communication design principles to enhance form, readability, and meaning.</span></span>

### <a name="interaction-design-principles"></a><span data-ttu-id="a8897-175">Principes de conception de l’interaction</span><span class="sxs-lookup"><span data-stu-id="a8897-175">Interaction design principles</span></span>

- <span data-ttu-id="a8897-176">Concevez votre projet de façon à permettre l’exploration.</span><span class="sxs-lookup"><span data-stu-id="a8897-176">Design to allow for exploration.</span></span>
- <span data-ttu-id="a8897-177">Autorisez les interactions directes avec des objets qui révèlent de nouvelles perspectives (tri par glissement, par exemple).</span><span class="sxs-lookup"><span data-stu-id="a8897-177">Allow for direct interactions with objects that reveal new insights (sorting via drag, for example).</span></span>
- <span data-ttu-id="a8897-178">Utilisez des modèles d’interaction simples, directs et familiers.</span><span class="sxs-lookup"><span data-stu-id="a8897-178">Use simple, direct, familiar interaction models.</span></span>

<span data-ttu-id="a8897-179">Pour plus d’informations sur la conception de visualisations de données interactives et conviviales, voir [Fondements et pièges de l’interface utilisateur](http://uitraps.com/).</span><span class="sxs-lookup"><span data-stu-id="a8897-179">For more information about how to design user-friendly interactive data visualizations, see [UI Tenets and Traps](http://uitraps.com/).</span></span>

### <a name="motion-design-principles"></a><span data-ttu-id="a8897-180">Principes de conception de mouvements</span><span class="sxs-lookup"><span data-stu-id="a8897-180">Motion design principles</span></span>

<span data-ttu-id="a8897-p119">Le mouvement suit un stimulus. Les éléments visuels doivent se déplacer dans la même direction à la même vitesse. Cela s’applique à :</span><span class="sxs-lookup"><span data-stu-id="a8897-p119">Motion follows stimulus. Visual elements should move in the same direction at the same rate. This applies to:</span></span>

- <span data-ttu-id="a8897-184">Création de graphique</span><span class="sxs-lookup"><span data-stu-id="a8897-184">Chart creation</span></span>
- <span data-ttu-id="a8897-185">Transition d’un type de graphique à un autre</span><span class="sxs-lookup"><span data-stu-id="a8897-185">Transition from one chart type to another chart type</span></span>
- <span data-ttu-id="a8897-186">Filtrage</span><span class="sxs-lookup"><span data-stu-id="a8897-186">Filtering</span></span>
- <span data-ttu-id="a8897-187">Tri</span><span class="sxs-lookup"><span data-stu-id="a8897-187">Sorting</span></span>
- <span data-ttu-id="a8897-188">Ajout ou soustraction de données</span><span class="sxs-lookup"><span data-stu-id="a8897-188">Adding or subtracting data</span></span>
- <span data-ttu-id="a8897-189">Brossage ou segmentation des données</span><span class="sxs-lookup"><span data-stu-id="a8897-189">Brushing or slicing data</span></span>
- <span data-ttu-id="a8897-190">Redimensionnement d’un graphique</span><span class="sxs-lookup"><span data-stu-id="a8897-190">Resizing a chart</span></span>

<span data-ttu-id="a8897-p120">Créez une perception de causalité. Lors de la préparation des animations :</span><span class="sxs-lookup"><span data-stu-id="a8897-p120">Create a perception of causality. When staging animations:</span></span>

- <span data-ttu-id="a8897-193">Préparez une chose à la fois.</span><span class="sxs-lookup"><span data-stu-id="a8897-193">Stage one thing at a time.</span></span> 
- <span data-ttu-id="a8897-194">Préparez les modifications des axes avant les modifications de l’encre de données.</span><span class="sxs-lookup"><span data-stu-id="a8897-194">Stage changes to axes before changes to data ink.</span></span>
- <span data-ttu-id="a8897-195">Préparez et animez des objets en tant que groupes s’ils se déplacent à la même vitesse dans la même direction.</span><span class="sxs-lookup"><span data-stu-id="a8897-195">Stage and animate objects as a group if they are moving at the same speed in the same direction.</span></span>
- <span data-ttu-id="a8897-p121">Préparez les éléments de données en groupes de 4 à 5 objets maximum. Les visionneuses ont des difficultés à suivre plus de 4 à 5 objets indépendamment.</span><span class="sxs-lookup"><span data-stu-id="a8897-p121">Stage data elements in groups of no more than 4-5 objects. Viewers have difficulty tracking more than 4-5 objects independently.</span></span>

<span data-ttu-id="a8897-198">Le mouvement ajoute une signification.</span><span class="sxs-lookup"><span data-stu-id="a8897-198">Motion adds meaning.</span></span>

- <span data-ttu-id="a8897-199">Les animations augmentent la compréhension par l’utilisateur des modifications apportées aux données, fournissent du contexte et agissent comme un calque d’annotation non verbal.</span><span class="sxs-lookup"><span data-stu-id="a8897-199">Animations increase user comprehension of changes to the data, provide context, and act as a non-verbal annotation layer.</span></span>
- <span data-ttu-id="a8897-200">Le mouvement doit se produire dans un espace de coordonnées significatif de la visualisation.</span><span class="sxs-lookup"><span data-stu-id="a8897-200">Motion should occur in a meaningful coordinate space of the visualization.</span></span>
- <span data-ttu-id="a8897-201">Adaptez l’animation en fonction du visuel.</span><span class="sxs-lookup"><span data-stu-id="a8897-201">Tailor the animation to the visual.</span></span> 
- <span data-ttu-id="a8897-202">Évitez les animations gratuites.</span><span class="sxs-lookup"><span data-stu-id="a8897-202">Avoid gratuitous animations.</span></span>

<span data-ttu-id="a8897-203">Le mouvement suit les données.</span><span class="sxs-lookup"><span data-stu-id="a8897-203">Motion follows data.</span></span>

- <span data-ttu-id="a8897-p122">Conservez les mappages de données. Si une zone est liée à une mesure, conservez cette zone de transition.</span><span class="sxs-lookup"><span data-stu-id="a8897-p122">Preserve data mappings. If an area is tied to a measure, maintain that area in transition.</span></span>
- <span data-ttu-id="a8897-p123">Maintenez un langage de création d’animation cohérent. Autant que possible, mappez l’animation de visualisation de données sur le langage de conception de mouvement Office existant. Utilisez des animations similaires pour les types de graphiques similaires.</span><span class="sxs-lookup"><span data-stu-id="a8897-p123">Maintain a consistent animation design language. Where possible, map data visualization animation to existing Office motion design language. Use similar animations for similar chart types.</span></span>

## <a name="accessibility-in-data-visualizations"></a><span data-ttu-id="a8897-209">Accessibilité des visualisations de données</span><span class="sxs-lookup"><span data-stu-id="a8897-209">Accessibility in data visualizations</span></span>

- <span data-ttu-id="a8897-p124">N’utilisez pas la couleur comme l’unique vecteur de communication des informations. Les personnes daltoniennes ne seront pas capables d’interpréter les résultats. Utilisez la forme, la taille et la texture en plus de la couleur lorsque cela est possible pour communiquer des informations.</span><span class="sxs-lookup"><span data-stu-id="a8897-p124">Do not use color as the only way to communicate information. People who are color blind will not be able to interpret the results. Use shape, size and texture in addition to color when possible to communicate information.</span></span>
- <span data-ttu-id="a8897-213">Rendez tous les éléments interactifs, tels que les boutons de commande ou les listes déroulantes, accessibles à partir d’un clavier.</span><span class="sxs-lookup"><span data-stu-id="a8897-213">Make all interactive elements, such as push buttons or pick lists, accessible from a keyboard.</span></span>
- <span data-ttu-id="a8897-214">Envoyez des événements d’accessibilité aux lecteurs d’écran pour annoncer les modifications de sélection, les info-bulles et ainsi de suite.</span><span class="sxs-lookup"><span data-stu-id="a8897-214">Send accessibility events to screen readers to announce focus changes, tooltips, and so on.</span></span>

## <a name="see-also"></a><span data-ttu-id="a8897-215">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a8897-215">See also</span></span> 

- [<span data-ttu-id="a8897-216">Données + Conception : une introduction simple à la préparation et à la visualisation des informations</span><span class="sxs-lookup"><span data-stu-id="a8897-216">Data + Design: A Simple Introduction to Preparing and Visualizing Information</span></span>](https://infoactive.co/data-design)
- [<span data-ttu-id="a8897-217">Les cinq meilleures bibliothèques pour créer des visualisations de données</span><span class="sxs-lookup"><span data-stu-id="a8897-217">The Five Best Libraries for Building Data Visualizations</span></span>](http://www.fastcompany.com/3029760/the-five-best-libraries-for-building-data-vizualizations)
- [<span data-ttu-id="a8897-218">Affichage visuel des informations quantitatives</span><span class="sxs-lookup"><span data-stu-id="a8897-218">The Visual Display of Quantitative Information</span></span>](https://www.edwardtufte.com/tufte/books_vdqi)
