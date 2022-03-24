---
title: Utilisation du mouvement dans les compléments Office
description: Obtenez les meilleures pratiques en matière d’utilisation des transitions, du mouvement ou de l’animation dans Office des modules.
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 3d5d67e50f3e1d2a0883c35bf65d23098dda6868
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743139"
---
# <a name="using-motion-in-office-add-ins"></a>Utilisation du mouvement dans les compléments Office

Lorsque vous concevez un complément Office, vous pouvez utiliser le mouvement pour améliorer l’expérience utilisateur. Les composants, contrôles et éléments de l’interface utilisateur ont souvent des comportements interactifs qui nécessitent des transitions, du mouvement ou de l’animation. Les caractéristiques de mouvement communes dans les éléments de l’interface utilisateur définissent les aspects d’animation d’un langage de création.

Office étant axé sur la productivité, le langage d’animation Office aide les clients dans l’exécution de leurs tâches. Il offre un équilibre entre réponse performante, chorégraphie fiable et satisfaction détaillée. Les compléments intégrés dans Office appartiennent à ce langage d’animation existant. Vu ce contexte, il est important de prendre en compte les recommandations suivantes lors de l’application d’un mouvement.

## <a name="create-motion-with-a-purpose"></a>Créer un mouvement avec un objectif

Le mouvement doit avoir un objectif qui transmet une valeur supplémentaire à l’utilisateur. Analysez le style et l’objectif de votre contenu lors du choix des animations. Gérez les messages critiques différemment des navigations d’exploration.

Les éléments standard utilisés dans un complément peuvent intégrer du mouvement permettant de se concentrer sur l’utilisateur, d’afficher les relations entre les éléments et de valider les actions de l’utilisateur. Chorégraphiez les éléments pour renforcer la hiérarchie et les modèles mentaux.

### <a name="best-practices"></a>Meilleures pratiques

|À faire|À ne pas faire|
|:-----|:-----|
|Identifiez les éléments clés dans le complément qui doivent avoir du mouvement. Les éléments le plus souvent animés dans un complément sont les panneaux, les superpositions, les fenêtres modales, les info-bulles, les menus et les légendes instructives.| Ne surchargez pas l’écran de l’utilisateur en animant tous les éléments. Évitez d’appliquer plusieurs mouvements visant à diriger ou guider l’utilisateur en attirant son attention sur de nombreux éléments en même temps. |
|Utilisez un mouvement simple et discret qui se comporte de manière attendue. Prenez en compte l’origine de votre élément déclencheur. Utilisez le mouvement pour créer un lien entre l’action et l’interface utilisateur obtenue. | Ne créez pas de temps d’attente pour un mouvement. Le mouvement dans les compléments ne doit pas altérer la fin de la tâche.|

![GIF qui affiche une ouverture de panneau avec des éléments de déplacement minimaux à côté d’une image GIF qui affiche une ouverture de panneau avec de nombreux éléments de déplacement.](../images/add-in-motion-purpose.gif)

## <a name="use-expected-motions"></a>Utiliser des mouvements attendus

Nous vous recommandons d’utiliser [l’interface](https://developer.microsoft.com/fluentui#/) utilisateur Fluent pour créer une connexion visuelle avec la plateforme Office, et nous encourageons également l’utilisation [d’animations](https://developer.microsoft.com/fluentui#/styles/web/motion) d’interface utilisateur Fluent pour créer des mouvements qui s’alignent sur le langage de mouvement Fabric.

Elle permet l’intégration en toute transparence dans Office. Elle vous aide à créer des expériences davantage ressenties qu’observées. Les classes CSS d’animation fournissent des informations de direction, d’entrée/sortie et de durée qui renforcent les modèles mentaux d’Office et offrent aux clients la possibilité d’apprendre à interagir avec votre complément.

### <a name="best-practices"></a>Meilleures pratiques

|À faire|À ne pas faire|
|:-----|:-----|
|Utilisez un mouvement qui s’aligne sur les comportements dans Fluent’interface utilisateur.| Ne créez pas de mouvements qui interfèrent ou entrent en conflit avec les modèles courants de mouvement dans Office.
|Assurez-vous qu’il existe une application cohérente du mouvement sur des éléments tels que.| N’utilisez pas de mouvements différents pour animer le même composant ou le même objet.|
|Assurez la cohérence de la direction dans l’animation. Par exemple, un panneau qui s’ouvre depuis le côté droit doit fermer vers le côté droit.|N’animez pas un élément en utilisant plusieurs directions.

![Gif qui affiche une ouverture modale d’une manière attendue à côté d’une image GIF qui affiche une ouverture modale d’une manière inattendue.](../images/add-in-motion-expected.gif)

## <a name="avoid-out-of-character-motion-for-an-element"></a>Éviter le mouvement de caractère pour les éléments

Prenez en compte la taille de la zone de dessin HTML (volet des tâches, boîte de dialogue ou complément de contenu) lors de l’implémentation du mouvement. Évitez de surcharger les espaces restreints. La mise en mouvement des éléments doit être compatible avec Office. Le caractère d’un mouvement de complément doit être performant, fiable et fluide. Au lieu d’entraver votre productivité, cherchez à informer et à diriger l’utilisateur.

### <a name="best-practices"></a>Meilleures pratiques

|À faire|À ne pas faire|
|:-----|:-----|
| Utilisez les [durées recommandées de mouvement](https://developer.microsoft.com/fluentui#/styles/web/motion). | N’utilisez pas trop d’animations. Évitez de créer des expériences qui enjolivent seulement l’interface utilisateur et détournent l’attention de vos clients.
| Suivez [les courbes d’accélération recommandées](/windows/uwp/design/motion/timing-and-easing#easing-in-fluent-motion).  |Ne mettez pas en mouvement les éléments de manière saccadée ou décousue. Évitez les anticipations, les rebonds, les élastiques ou autres effets qui émulent la physique du monde naturel.|

![GIF qui affiche le chargement des vignettes à l’aide d’une fondue légère à côté d’une image GIF qui affiche le chargement des vignettes avec rebond.](../images/add-in-motion-character.gif)

## <a name="see-also"></a>Voir aussi

* [Fluent’animation de l’interface utilisateur](https://developer.microsoft.com/fluentui#/styles/web/motion)
* [Mouvement pour les applications de la plateforme Windows universelle](/windows/uwp/design/motion)
