# <a name="office-add-in-design-language"></a>Langage de création d’un complément Office

Le langage de création d’Office est un système visuel clair et simple qui garantit la cohérence entre expériences. Il contient un ensemble d’éléments visuels qui définissent les interfaces Office, y compris : 

- Police standard
- Palette de couleurs courantes
- Ensemble de tailles typographiques et pondérations
- Instructions relatives aux icônes
- Éléments d’icône partagée
- Définitions d’animation
- Composants courants

[Office UI Fabric](https://dev.office.com/fabric) est l’infrastructure frontale officielle pour la création avec le langage de création Office. L’utilisation de Fabric est facultative, mais elle est le moyen le plus rapide pour vous assurer que vos compléments sont une extension naturelle d’Office. Profitez de Fabric pour concevoir et créer des compléments qui complètent Office.

De nombreux compléments d’Office sont associés à une marque préexistante. Vous pouvez conserver une marque forte et son langage de composant ou visuel dans votre complément. Recherchez les opportunités pour conserver votre propre langage visuel lors de l’intégration avec Office. Pensez à des moyens de remplacer les couleurs Office, la typographie, les icônes ou d’autres éléments stylistiques par des éléments de votre marque. Pensez à des moyens de suivre des dispositions de complément ou des modèles de conception de l’expérience utilisateur courants tout en insérant des contrôles et des composants que vos clients connaissent.

L’insertion d’une interface utilisateur HTML de marque importante à l’intérieur d’Office peut créer des dissonances pour les clients. Trouvez un équilibre qui s’adapte en toute transparence dans Office mais qui s’aligne aussi clairement sur votre marque parent ou de service. Lorsqu’un complément ne s’adapte pas à Office, c’est souvent en raison d’une incompatibilité des éléments stylistiques. Par exemple, la typographie est trop grande et en dehors de la grille, les couleurs sont particulièrement criardes ou contrastées, ou les animations sont superflues et se comportent différemment par rapport à Office. L’apparence et le comportement des contrôles ou des composants dévient trop des normes d’Office.

## <a name="typography"></a>Typographie
Segoe est la police standard pour Office. Utilisez-la dans votre complément pour être en adéquation avec les volets des tâches, les boîtes de dialogue et les objets de contenu d’Office. Office UI Fabric vous donne accès à Segoe. Il fournit un dégradé de polices complet de Segoe avec de nombreuses variations, d’épaisseur de police et de taille, dans des classes CSS pratiques. Toutes les tailles et épaisseurs de police d’Office UI Fabric n’ont pas une belle apparence dans un complément Office. Pour une intégration harmonieuse ou pour éviter les incompatibilités, envisagez d’utiliser un sous-ensemble du dégradé de polices de Fabric. Voici une liste des classes de base de la structure que nous vous recommandons d’utiliser dans les compléments Office.

|Exemple |Classe |Taille |Pondération |Utilisation recommandée |
|------ |----- |---- |------ |----------------- |
|![Image de texte Hero](../../images/add-in-typeramp-hero.png)|.ms-font-xxl |28 px | Segoe Light |<ul><li>Cette classe est plus grande que tous les autres éléments typographiques dans Office. Utilisez-la avec parcimonie pour éviter une hiérarchie visuelle non valide.</li><li>Évitez d’utiliser de longues chaînes dans des espaces limités.</li><li>Laissez suffisamment d’espaces blancs autour du texte en utilisant cette classe.</li><li>Couramment utilisée pour les premiers messages, éléments hero ou autres appels à l’action.</li></ul> |
|![Image de texte Hero](../../images/add-in-typeramp-title.png)|.ms-font-xl |21 px |Segoe Light | <ul><li>Cette classe correspond au titre du volet des tâches des applications Office.</li><li>Utilisez-la avec parcimonie pour éviter une hiérarchie typographique plate.</li><li>Couramment utilisée comme élément de niveau supérieur (titres de contenu, de page ou de boîte de dialogue).</li><li></ul> |
|![Image de texte Hero](../../images/add-in-typeramp-subtitle.png)|.ms-font-l |17 px |Segoe Semilight | <ul><li>Cette classe est le premier point en dessous des titres.</li><li>Couramment utilisée comme sous-titre, élément de navigation ou en-tête de groupe.</li><ul> |
|![Image de texte Hero](../../images/add-in-typeramp-body.png)|.ms-font-m |14 px |Segoe Regular |* Couramment utilisée comme corps de texte dans les compléments. |
|![Image de texte Hero](../../images/add-in-typeramp-caption.png)|.ms-font-xs |11 px | Segoe Regular |* Couramment utilisée pour le texte secondaire ou tertiaire (horodateurs, par lignes, légendes ou étiquettes de champ). |
|![Image de texte Hero](../../images/add-in-typeramp-annotation.png)|.ms-font-mi |10 px |Segoe Semibold |* Le plus petit niveau dans le dégradé de polices doit être rarement utilisé. Il est disponible lorsque la lisibilité n’est pas requise. |
> La couleur du texte n’est pas incluse dans ces classes de base. Utilisez « Neutre primaire » de Fabric pour la plupart du texte sur arrière-plans blancs.

## <a name="color"></a>Couleur
La couleur est souvent utilisée pour mettre en évidence la marque et renforcer la hiérarchie visuelle. Elle permet d’identifier une interface et de guider les clients dans une expérience. Dans Office, la couleur est utilisée pour les mêmes objectifs mais elle est appliquée délibérément et au minimum. Elle ne surcharge jamais le contenu clients. Même lorsque chaque application Office est marquée avec sa propre couleur dominante, elle est utilisée avec parcimonie.

Office UI Fabric comprend un jeu de couleurs de thème par défaut. Lorsque Fabric est appliqué à un complément Office comme composants ou dans des dispositions, les mêmes objectifs s’appliquent. La couleur doit communiquer la hiérarchie, guidant ainsi les clients vers l’action sans interférer avec le contenu. Les couleurs de thème Fabric peuvent introduire une nouvelle couleur de l’accentuation dans l’interface globale. Cette nouvelle accentuation peut entrer en conflit avec la personnalisation de l’application Office et interférer avec la hiérarchie. En d’autres termes, Fabric peut introduire une nouvelle couleur de l’accentuation dans l’interface globale lorsqu’elle est utilisée à l’intérieur d’un complément. Cette nouvelle couleur de l’accentuation peut créer une confusion et interférer avec la hiérarchie globale. Envisagez des façons d’éviter les conflits et les interférences. Utilisez des accentuations neutres ou remplacez les couleurs de thème Fabric en fonction de la personnalisation de l’application Office ou de vos propres couleurs de la marque.

Les applications Office permettent aux clients de personnaliser leurs interfaces en appliquant un thème de l’interface utilisateur d’Office. Les clients peuvent choisir entre quatre thèmes de l’interface utilisateur pour modifier le style des arrière-plans et des boutons dans Word, PowerPoint, Excel et les autres applications de la suite Office. Pour que vos compléments paraissent comme des composants naturels d’Office et répondent à la personnalisation, utilisez nos API de thèmes. Par exemple, les couleurs d’arrière-plan du volet des tâches deviennent gris foncé dans certains thèmes. Nos API de thèmes vous permettent de faire de même et d’ajuster le texte de premier plan pour garantir l’[accessibilité](add-in-design-guidelines.md#accessibility-guidelines).

>  Pour les compléments de volet de tâches et de messagerie, utilisez la propriété [Context.officeTheme](https://dev.office.com/docs/reference/shared/office.context.officetheme.htm) pour utiliser les thèmes correspondant à ceux des applications Office. Actuellement, cette API n’est disponible que dans Office 2016.

> Pour plus d’informations sur les compléments de contenu pour PowerPoint, voir [Utiliser des thèmes Office dans vos compléments PowerPoint](https://dev.office.com/docs/add-ins/powerpoint/use-document-themes-in-your-powerpoint-add-ins.htm).

Appliquez les recommandations générales suivantes pour la couleur :

* Utilisez la couleur avec parcimonie pour communiquer la hiérarchie et renforcer la marque.
* L’utilisation excessive d’une couleur d’accentuation unique appliquée aux éléments interactifs et non interactifs peut être source de confusion. Par exemple, évitez d’utiliser la même couleur pour les éléments sélectionnés et non sélectionnés dans un menu de navigation.
* Évitez les conflits inutiles avec des couleurs non Office.
* Utilisez vos propres couleurs de la marque pour créer une association avec votre service ou votre société.
* Assurez-vous que tout le texte est accessible. N’oubliez pas qu’il existe un ratio de contraste 4.5:1 entre le texte de premier plan et l’arrière-plan.
* Pensez au daltonisme. Utilisez plus que simplement de la couleur pour indiquer l’interactivité et la hiérarchie.
* Consultez [Instructions relatives aux icônes](design-icons.md) pour en savoir plus sur la conception des icônes de commande de complément avec la palette de couleurs d’icônes Office.

## <a name="layout"></a>Disposition
Chaque conteneur HTML incorporé dans Office aura une disposition. Ces dispositions sont les écrans principaux de votre complément. Dans ces dispositions, vous créerez des expériences qui permettent aux clients de lancer des actions, de modifier des paramètres, d’afficher, de faire défiler ou de parcourir du contenu. Concevez votre complément avec une disposition cohérente à travers les écrans afin de garantir la continuité de l’expérience. Si vous avez un site web existant que vos clients utilisent souvent, envisagez de réutiliser les dispositions de vos pages web existantes. Adaptez-les pour qu’elles s’intègrent harmonieusement dans des conteneurs HTML Office.

Pour des recommandations sur la disposition, voir [Volet des tâches](task-pane-add-ins.md), [Contenu](content-add-ins.md) et [Boîte de dialogue](dialog-boxes.md). Pour plus d’informations sur la façon d’assembler des composants Office UI Fabric dans des flux d’expérience utilisateur et des dispositions courants , voir [Modèles de conception UX](ux-design-patterns.md).

Appliquez les recommandations générales suivantes pour les dispositions :

*    Évitez les marges étroites ou larges sur vos conteneurs HTML. 20 pixels est une grande valeur par défaut. 
*    Alignez les éléments intentionnellement. Les retraits supplémentaires et les nouveaux points d’alignement doivent aider la hiérarchie visuelle.
*    Les interfaces Office se trouvent sur une grille 4px. Essayez de conserver votre marge intérieure entre les éléments à des multiples de 4. 
*    Une interface surchargée peut être source de confusion et ne pas être utilisée facilement avec les interactions tactiles. 
*    Vérifiez que les dispositions sont cohérentes entre les écrans. Les modifications de disposition inattendues ressemblent à des bogues visuels qui contribuent à un manque de confiance en votre solution. 
*    Suivez les modèles de disposition courants. Les conventions permettent aux utilisateurs de comprendre comment utiliser une interface.
*    Évitez les éléments redondants comme la personnalisation ou les commandes.
*    Consolidez les contrôles et les affichages pour éviter une utilisation excessive de la souris. 
*    Créez des expériences réactives qui s’adaptent aux hauteurs et largeurs du conteneur HTML.

## <a name="component-language"></a>Langage du composant

Les écrans et les dispositions sont constitués de contenu et de composants. Les composants sont des contrôles qui aident vos clients à interagir avec les éléments de votre logiciel ou service. Les boutons, la navigation, les badges, les alertes et les menus déroulants sont tous des exemples de composants courants qui ont souvent des comportements et des styles cohérents.

Office UI Fabric rend les composants qui ressemblent à une partie d’Office et se comportent comme une partie d’Office. Utilisez Fabric pour l’intégration transparente avec Office. Si votre complément a son propre langage de composant préexistant, vous n’avez pas besoin de l’abandonner en faveur de Fabric. Recherchez les opportunités pour le conserver lors de l’intégration avec Office. Pensez à remplacer les éléments stylistiques, à supprimer les conflits ou à adopter des styles et des comportements qui éliminent la confusion de l’utilisateur.

Appliquez les recommandations générales suivantes pour les composants :

*    Ne répliquez pas le ruban Office à l’intérieur de votre complément
*    Évitez de créer des menus, des boutons ou d’autres composants qui se comportent différemment des composants Office.
*    Utilisez les composants [Office UI Fabric](office-ui-fabric.md) que nous recommandons pour les compléments.
*    Utilisez les [modèles de conception UX](ux-design-patterns.md) pour les composants de l’interface utilisateur d’Office courants. 

## <a name="icons"></a>Icônes
Les icônes sont la représentation visuelle d’un comportement ou d’un concept. Elles sont souvent utilisées pour ajouter une signification aux contrôles et commandes. Les visuels, qu’ils soient réalistes ou symboliques, permettent à l’utilisateur de naviguer dans l’interface utilisateur de la même façon que les signes l’aident à naviguer dans son environnement. Ils doivent être simples et clairs et contenir uniquement les informations nécessaires pour permettre aux clients d’analyser rapidement l’action qui se produit lorsqu’ils choisissent un contrôle.

Les interfaces de ruban Office ont un style visuel standard. Si vous concevez une commande de complément pour le ruban Office, suivez nos [instructions relatives aux icônes](design-icons). Cela garantit la cohérence dans les applications Office. Les instructions vous aident à créer un ensemble de composants PNG pour votre solution qui s’intègrent naturellement dans Office.

De nombreux conteneurs HTML contiennent des contrôles avec iconographie. Utilisez la police personnalisée d’Office UI Fabric pour le rendu des icônes de style Office dans votre complément. La police d’icône de Fabric contient de nombreux glyphes pour les métaphores Office courantes que vous pouvez redimensionner, colorier et personnaliser selon vos besoins. Si vous avez un langage visuel existant avec votre propre jeu d’icônes, n’hésitez pas à l’utiliser dans vos canevas HTML. Créer la continuité avec votre marque avec un jeu d’icônes standard est une partie importante de tout langage de création. Soyez prudent pour éviter de créer de la confusion pour les clients en conflit avec les métaphores Office.

Appliquez les recommandations générales suivantes pour les icônes :

* Ne redéfinissez pas les glyphes Office UI Fabric pour les commandes de complément dans le ruban Office ou les menus contextuels. Les icônes Fabric sont stylistiquement différentes et ne correspondront pas.
* Utilisez le langage d’icône Office pour représenter des comportements ou des concepts.
* Réutilisez les métaphores visuelles d’Office courantes telles que le pinceau pour mettre en forme ou la loupe pour rechercher.
* N’utilisez pas les métaphores pour des actions qui n’ont rien à voir. L’utilisation du même visuel pour un comportement ou un concept différent peut être source de confusion pour les utilisateurs.

## <a name="animation"></a>Animation
Les composants, contrôles et éléments de l’interface utilisateur ont souvent des comportements interactifs qui nécessitent des transitions, du mouvement ou de l’animation. Les caractéristiques de mouvement communes dans les éléments de l’interface utilisateur définissent les aspects d’animation d’un langage de création. Office étant axé sur la productivité, le langage d’animation Office aide les clients dans l’exécution de leurs tâches. Il offre un équilibre entre réponse performante, chorégraphie fiable et satisfaction détaillée.

Office UI Fabric comprend une bibliothèque d’animation pour contrôler le mouvement dans vos conteneurs HTML. Elle permet l’intégration en toute transparence dans Office. Elle vous aide à créer des expériences davantage ressenties qu’observées. Les classes CSS d’animation fournissent des informations de direction, entrée/sortie et durée qui renforcent les modèles mentaux d’Office et offrent aux clients la possibilité d’apprendre à interagir avec votre complément. 

Si votre complément a son propre langage d’animation, utilisez-le. Recherchez les opportunités pour conserver votre animation marquée lors de l’intégration avec Office. Veillez à ne pas interférer ni à entrer en conflit avec les modèles de mouvement courants dans Office. Évitez de créer des expériences qui sont embellissements qui ne font que créer une confusion pour vos clients.

Appliquez les recommandations générales suivantes pour les animations :

* Les animations doivent être ressenties et vécues inconsciemment, afin d’éviter d’altérer la fin de la tâche.
* Évitez les anticipations, les rebonds, les élastiques ou autres effets qui émulent la physique du monde naturel.
* Chorégraphiez les éléments pour renforcer la hiérarchie et les modèles mentaux.
* Utilisez le mouvement pour guider l’utilisateur et fournir un focus de composition sur les éléments clés pour l’exécution d’une tâche. 
* Pensez à l’origine de votre élément déclencheur. Utilisez le mouvement pour créer un lien entre l’action et l’interface utilisateur obtenue.
* Analysez le style et l’objectif de votre contenu lors du choix des animations. Gérez les messages critiques différemment des navigations d’exploration.
