---
title: Langage de cr?ation d?un compl?ment Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 7d19714fa14fb374bcd41aa744c08929c228c94f
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="office-add-in-design-language"></a>Langage de cr?ation d?un compl?ment Office

Le langage de cr?ation d?Office est un syst?me visuel clair et simple qui garantit la coh?rence entre exp?riences. Il contient un ensemble d??l?ments visuels qui d?finissent les interfaces Office, y compris :

- Police standard
- Palette de couleurs courantes
- Ensemble de tailles typographiques et pond?rations
- Instructions relatives aux ic?nes
- ?l?ments d?ic?ne partag?e
- D?finitions d?animation
- Composants courants

[Office UI Fabric](https://dev.office.com/fabric) est l?infrastructure frontale officielle pour la cr?ation avec le langage de cr?ation Office. L?utilisation de Fabric est facultative, mais elle est le moyen le plus rapide pour vous assurer que vos compl?ments sont une extension naturelle d?Office. Profitez de Fabric pour concevoir et cr?er des compl?ments qui compl?tent Office.

De nombreux compl?ments d?Office sont associ?s ? une marque pr?existante. Vous pouvez conserver une marque forte et son langage de composant ou visuel dans votre compl?ment. Recherchez les opportunit?s pour conserver votre propre langage visuel lors de l?int?gration avec Office. Pensez ? des moyens de remplacer les couleurs Office, la typographie, les ic?nes ou d?autres ?l?ments stylistiques par des ?l?ments de votre marque. Pensez ? des moyens de suivre des dispositions de compl?ment ou des mod?les de conception de l?exp?rience utilisateur courants tout en ins?rant des contr?les et des composants que vos clients connaissent.

L?insertion d?une interface utilisateur HTML de marque importante ? l?int?rieur d?Office peut cr?er des dissonances pour les clients. Trouvez un ?quilibre qui s?adapte en toute transparence dans Office mais qui s?aligne aussi clairement sur votre marque parent ou de service. Lorsqu?un compl?ment ne s?adapte pas ? Office, c?est souvent en raison d?une incompatibilit? des ?l?ments stylistiques. Par exemple, la typographie est trop grande et en dehors de la grille, les couleurs sont particuli?rement criardes ou contrast?es, ou les animations sont superflues et se comportent diff?remment par rapport ? Office. L?apparence et le comportement des contr?les ou des composants d?vient trop des normes d?Office.

## <a name="typography"></a>Typographie
Segoe est la police standard pour Office. Utilisez-la dans votre compl?ment pour ?tre en ad?quation avec les volets des t?ches, les bo?tes de dialogue et les objets de contenu d?Office. Office UI Fabric vous donne acc?s ? Segoe. Il fournit un d?grad? de polices complet de Segoe avec de nombreuses variations, d??paisseur de police et de taille, dans des classes CSS pratiques. Toutes les tailles et ?paisseurs de police d?Office UI Fabric n?ont pas une belle apparence dans un compl?ment Office. Pour une int?gration harmonieuse ou pour ?viter les incompatibilit?s, envisagez d?utiliser un sous-ensemble du d?grad? de polices de Fabric. Voici une liste des classes de base de la structure que nous vous recommandons d?utiliser dans les compl?ments Office.

|Exemple |Classe |Taille |Pond?ration |Utilisation recommand?e |
|------ |----- |---- |------ |----------------- |
|![Image de texte Hero](../images/add-in-typeramp-hero.png)|.ms-font-xxl |28 px | Segoe Light |<ul><li>Cette classe est plus grande que tous les autres ?l?ments typographiques dans Office. Utilisez-la avec parcimonie pour ?viter une hi?rarchie visuelle non valide.</li><li>?vitez d?utiliser de longues cha?nes dans des espaces limit?s.</li><li>Laissez suffisamment d?espaces blancs autour du texte en utilisant cette classe.</li><li>Couramment utilis?e pour les premiers messages, ?l?ments hero ou autres appels ? l?action.</li></ul> |
|![Image de texte Hero](../images/add-in-typeramp-title.png)|.ms-font-xl |21 px |Segoe Light | <ul><li>Cette classe correspond au titre du volet des t?ches des applications Office.</li><li>Utilisez-la avec parcimonie pour ?viter une hi?rarchie typographique plate.</li><li>Couramment utilis?e comme ?l?ment de niveau sup?rieur (titres de contenu, de page ou de bo?te de dialogue).</li></ul> |
|![Image de texte Hero](../images/add-in-typeramp-subtitle.png)|.ms-font-l |17 px |Segoe Semilight | <ul><li>Cette classe est le premier point en dessous des titres.</li><li>Couramment utilis?e comme sous-titre, ?l?ment de navigation ou en-t?te de groupe.</li><ul> |
|![Image de texte Hero](../images/add-in-typeramp-body.png)|.ms-font-m |14 px |Segoe Regular |<ul><li>Couramment utilis?e comme corps de texte dans les compl?ments.</li><ul>|
|![Image de texte Hero](../images/add-in-typeramp-caption.png)|.ms-font-xs |11 px | Segoe Regular |<ul><li>Couramment utilis?e pour le texte secondaire ou tertiaire (horodatages, signatures, l?gendes ou ?tiquettes de champ).</li><ul>|
|![Image de texte Hero](../images/add-in-typeramp-annotation.png)|.ms-font-mi |10 px |Segoe Semibold |<ul><li>Le plus petit niveau dans le d?grad? de polices doit ?tre rarement utilis?. Il est disponible lorsque la lisibilit? n?est pas requise.</li><ul>|

> [!NOTE]
> La couleur du texte n?est pas incluse dans ces classes de base. Utilisez le ? neutre primaire ? de Fabric pour la plupart du texte sur des arri?re-plans blancs.

## <a name="color"></a>Couleur
La couleur est souvent utilis?e pour mettre en ?vidence l'identit? graphique et renforcer la hi?rarchie de l'objet visuel. Elle aide ? identifier une interface et ? guider les clients ? travers une exp?rience. Dans Office, la couleur est utilis?e pour les m?mes objectifs mais elle est appliqu?e de mani?re cibl?e et minimale. ? aucun moment, cela ne surcharge le contenu du client. M?me lorsque chaque application Office est marqu?e avec sa propre couleur dominante, elle est utilis?e avec parcimonie.

Office UI Fabric comprend un jeu de couleurs de th?me par d?faut. Lorsque Fabric est appliqu? ? un compl?ment Office comme composants ou dans des dispositions, les m?mes objectifs s?appliquent. La couleur doit communiquer la hi?rarchie, guidant ainsi les clients vers l?action sans interf?rer avec le contenu. Les couleurs de th?me Fabric peuvent introduire une nouvelle couleur de l?accentuation dans l?interface globale. Cette nouvelle accentuation peut entrer en conflit avec la personnalisation de l?application Office et interf?rer avec la hi?rarchie. En d?autres termes, Fabric peut introduire une nouvelle couleur de l?accentuation dans l?interface globale lorsqu?elle est utilis?e ? l?int?rieur d?un compl?ment. Cette nouvelle couleur de l?accentuation peut cr?er une confusion et interf?rer avec la hi?rarchie globale. Envisagez des fa?ons d??viter les conflits et les interf?rences. Utilisez des accentuations neutres ou remplacez les couleurs de th?me Fabric en fonction de la personnalisation de l?application Office ou de vos propres couleurs de la marque.

Les applications Office permettent aux clients de personnaliser leurs interfaces en appliquant un th?me de l?interface utilisateur d?Office. Les clients peuvent choisir entre quatre th?mes de l?interface utilisateur pour modifier le style des arri?re-plans et des boutons dans Word, PowerPoint, Excel et les autres applications de la suite Office. Pour que vos compl?ments paraissent comme des composants naturels d?Office et r?pondent ? la personnalisation, utilisez nos API de th?mes. Par exemple, les couleurs d?arri?re-plan du volet des t?ches deviennent gris fonc? dans certains th?mes. Nos API de th?mes vous permettent de faire de m?me et d?ajuster le texte de premier plan pour garantir l?[accessibilit?](add-in-design-guidelines.md#accessibility-guidelines).

> [!NOTE]
> - Pour les compl?ments de volet de t?ches et de messagerie, utilisez la propri?t? [Context.officeTheme](https://dev.office.com/reference/add-ins/shared/office.context.officetheme) pour utiliser les th?mes correspondant ? ceux des applications Office. Actuellement, cette API n?est disponible que dans Office 2016.
> - Pour plus d?informations sur les compl?ments de contenu pour PowerPoint, reportez-vous ? l?article expliquant comment [utiliser des th?mes Office dans vos compl?ments PowerPoint](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md).

Appliquez les recommandations g?n?rales suivantes pour la couleur :

* Utilisez la couleur avec parcimonie pour communiquer la hi?rarchie et renforcer la marque.
* L?utilisation excessive d?une couleur d?accentuation unique appliqu?e aux ?l?ments interactifs et non interactifs peut ?tre source de confusion. Par exemple, ?vitez d?utiliser la m?me couleur pour les ?l?ments s?lectionn?s et non s?lectionn?s dans un menu de navigation.
* ?vitez les conflits inutiles avec des couleurs non Office.
* Utilisez vos propres couleurs de la marque pour cr?er une association avec votre service ou votre soci?t?.
* Assurez-vous que tout le texte est accessible. N?oubliez pas qu?il existe un ratio de contraste 4.5:1 entre le texte de premier plan et l?arri?re-plan.
* Pensez aux personnes atteintes de daltonisme : n?utilisez pas que des couleurs pour indiquer l?interactivit? et la hi?rarchie.
* Consultez [Instructions relatives aux ic?nes](design-icons.md) pour en savoir plus sur la conception des ic?nes de commande de compl?ment avec la palette de couleurs d?ic?nes Office.

## <a name="layout"></a>Disposition
Chaque conteneur HTML incorpor? dans Office aura une disposition. Ces dispositions sont les ?crans principaux de votre compl?ment. Dans ces dispositions, vous cr?erez des exp?riences qui permettent aux clients de lancer des actions, de modifier des param?tres, d?afficher, de faire d?filer ou de parcourir du contenu. Concevez votre compl?ment avec une disposition coh?rente ? travers les ?crans afin de garantir la continuit? de l?exp?rience. Si vous avez un site web existant que vos clients utilisent souvent, envisagez de r?utiliser les dispositions de vos pages web existantes. Adaptez-les pour qu?elles s?int?grent harmonieusement dans des conteneurs HTML Office.

Pour des recommandations sur la disposition, voir [Volet des t?ches](task-pane-add-ins.md), [Contenu](content-add-ins.md) et [Bo?te de dialogue](dialog-boxes.md). Pour plus d?informations sur la fa?on d?assembler des composants Office UI Fabric dans des flux d?exp?rience utilisateur et des dispositions courants , voir [Mod?les de conception UX](ux-design-patterns.md).

Appliquez les recommandations g?n?rales suivantes pour les dispositions :

*   ?vitez les marges ?troites ou larges sur vos conteneurs HTML. 20 pixels est une grande valeur par d?faut.
*   Alignez les ?l?ments intentionnellement. Les retraits suppl?mentaires et les nouveaux points d?alignement doivent aider la hi?rarchie visuelle.
*   Les interfaces Office se trouvent sur une grille 4px. Essayez de conserver votre marge int?rieure entre les ?l?ments ? des multiples de 4.
*   Une interface surcharg?e peut ?tre source de confusion et ne pas ?tre utilis?e facilement avec les interactions tactiles.
*   V?rifiez que les dispositions sont coh?rentes entre les ?crans. Les modifications de disposition inattendues ressemblent ? des bogues visuels qui contribuent ? un manque de confiance en votre solution.
*   Suivez les mod?les de disposition courants. Les conventions permettent aux utilisateurs de comprendre comment utiliser une interface.
*   ?vitez les ?l?ments redondants comme la personnalisation ou les commandes.
*   Consolidez les contr?les et les affichages pour ?viter une utilisation excessive de la souris.
*   Cr?ez des exp?riences r?actives qui s?adaptent aux hauteurs et largeurs du conteneur HTML.

## <a name="component-language"></a>Langage du composant

Les ?crans et les dispositions sont constitu?s de contenu et de composants. Les composants sont des contr?les qui aident vos clients ? interagir avec les ?l?ments de votre logiciel ou service. Les boutons, la navigation, les badges, les alertes et les menus d?roulants sont tous des exemples de composants courants qui ont souvent des comportements et des styles coh?rents.

Office UI Fabric rend les composants qui ressemblent ? une partie d?Office et se comportent comme une partie d?Office. Utilisez Fabric pour l?int?gration transparente avec Office. Si votre compl?ment a son propre langage de composant pr?existant, vous n?avez pas besoin de l?abandonner en faveur de Fabric. Recherchez les opportunit?s pour le conserver lors de l?int?gration avec Office. Pensez ? remplacer les ?l?ments stylistiques, ? supprimer les conflits ou ? adopter des styles et des comportements qui ?liminent la confusion de l?utilisateur.

Appliquez les recommandations g?n?rales suivantes pour les composants :

*   Ne r?pliquez pas le ruban Office ? l?int?rieur de votre compl?ment
*   ?vitez de cr?er des menus, des boutons ou d?autres composants qui se comportent diff?remment des composants Office.
*   Utilisez les composants [Office UI Fabric](office-ui-fabric.md) que nous recommandons pour les compl?ments.
*   Utilisez les [mod?les de conception UX](ux-design-patterns.md) pour les composants de l?interface utilisateur d?Office courants.

## <a name="icons"></a>Ic?nes
Les ic?nes sont la repr?sentation visuelle d?un comportement ou d?un concept. Elles sont souvent utilis?es pour ajouter une signification aux contr?les et commandes. Les visuels, qu?ils soient r?alistes ou symboliques, permettent ? l?utilisateur de naviguer dans l?interface utilisateur de la m?me fa?on que les signes l?aident ? naviguer dans son environnement. Ils doivent ?tre simples et clairs et contenir uniquement les informations n?cessaires pour permettre aux clients d?analyser rapidement l?action qui se produit lorsqu?ils choisissent un contr?le.

Les interfaces de ruban Office ont un style visuel standard. Si vous concevez une commande de compl?ment pour le ruban Office, suivez nos [instructions relatives aux ic?nes](design-icons.md). Cela garantit la coh?rence dans les applications Office. Les instructions vous aident ? cr?er un ensemble de composants PNG pour votre solution qui s?int?grent naturellement dans Office.

De nombreux conteneurs HTML contiennent des contr?les avec iconographie. Utilisez la police personnalis?e d?Office UI Fabric pour le rendu des ic?nes de style Office dans votre compl?ment. La police d?ic?ne de Fabric contient de nombreux glyphes pour les m?taphores Office courantes que vous pouvez redimensionner, colorier et personnaliser selon vos besoins. Si vous avez un langage visuel existant avec votre propre jeu d?ic?nes, n?h?sitez pas ? l?utiliser dans vos canevas HTML. Cr?er la continuit? avec votre marque avec un jeu d?ic?nes standard est une partie importante de tout langage de cr?ation. Soyez prudent pour ?viter de cr?er de la confusion pour les clients en conflit avec les m?taphores Office.

Appliquez les recommandations g?n?rales suivantes pour les ic?nes :

* Ne red?finissez pas les glyphes Office UI Fabric pour les commandes de compl?ment dans le ruban Office ou les menus contextuels. Les ic?nes Fabric sont stylistiquement diff?rentes et ne correspondront pas.
* Utilisez le langage d?ic?ne Office pour repr?senter des comportements ou des concepts.
* R?utilisez les m?taphores visuelles d?Office courantes telles que le pinceau pour mettre en forme ou la loupe pour rechercher.
* N?utilisez pas les m?taphores pour des actions qui n?ont rien ? voir. L?utilisation du m?me visuel pour un comportement ou un concept diff?rent peut ?tre source de confusion pour les utilisateurs.


## <a name="see-also"></a>Voir aussi

- [Instructions de cr?ation d?un compl?ment Office](add-in-design-guidelines.md)
- [Utilisation du mouvement dans les compl?ments Office](using-motion-office-addins.md)
