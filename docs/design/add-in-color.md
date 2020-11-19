---
title: Instructions concernant les couleurs pour les compléments Office
description: Découvrez comment utiliser les couleurs dans l’interface utilisateur d’un complément Office.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 99eef66ec5ed1cb421d4d8cef7e20d8b19a0ee3d
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132185"
---
# <a name="color"></a>Couleur

La couleur est souvent utilisée pour mettre en évidence la marque et renforcer la hiérarchie visuelle. Elle permet d’identifier une interface et de guider les clients dans une expérience. Dans Office, la couleur est utilisée pour les mêmes objectifs mais elle est appliquée délibérément et au minimum. Elle ne surcharge jamais le contenu clients. Même lorsque chaque application Office est marquée avec sa propre couleur dominante, elle est utilisée avec parcimonie.

![Diagramme illustrant le jeu de couleurs pour Office, Excel, Word et PowerPoint. Les couleurs principales d’Office sont le noir et blanc, et les couleurs secondaires sont gris clair, gris foncé et orange. La couleur dominante d’Excel est le vert, Word est bleu et PowerPoint est orange.](../images/office-addins-color-schemes.png)

Office UI Fabric comprend un jeu de couleurs de thème par défaut. Lorsque Fabric est appliqué à un complément Office comme composants ou dans des dispositions, les mêmes objectifs s’appliquent. La couleur doit communiquer la hiérarchie, guidant ainsi les clients vers l’action sans interférer avec le contenu. Les couleurs de thème Fabric peuvent introduire une nouvelle couleur de l’accentuation dans l’interface globale. Cette nouvelle accentuation peut entrer en conflit avec la personnalisation de l’application Office et interférer avec la hiérarchie. En d’autres termes, Fabric peut introduire une nouvelle couleur de l’accentuation dans l’interface globale lorsqu’elle est utilisée à l’intérieur d’un complément. Cette nouvelle couleur de l’accentuation peut créer une confusion et interférer avec la hiérarchie globale. Envisagez des façons d’éviter les conflits et les interférences. Utilisez des accentuations neutres ou remplacez les couleurs de thème Fabric en fonction de la personnalisation de l’application Office ou de vos propres couleurs de la marque.

Les applications Office permettent aux clients de personnaliser leurs interfaces en appliquant un thème de l’interface utilisateur d’Office. Les clients peuvent choisir entre quatre thèmes de l’interface utilisateur pour modifier le style des arrière-plans et des boutons dans Word, PowerPoint, Excel et les autres applications de la suite Office. Pour que vos compléments ressemblent à une partie naturelle d’Office et répondent à la personnalisation, utilisez nos API de thèmes. Par exemple, les couleurs d’arrière-plan du volet des tâches deviennent gris foncé dans certains thèmes. Nos API de thèmes vous permettent de faire de même et d’ajuster le texte de premier plan pour garantir l’[accessibilité](../design/accessibility-guidelines.md).

> [!NOTE]
> - Pour les compléments de volet de tâches et de messagerie, utilisez la propriété [Context.officeTheme](/javascript/api/office/office.context) pour utiliser les thèmes correspondant à ceux des applications Office. Cette API est actuellement disponible dans Office 2016 ou version ultérieure.
> - Pour plus d’informations sur les compléments de contenu pour PowerPoint, reportez-vous à l’article expliquant comment [utiliser des thèmes Office dans vos compléments PowerPoint](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md).

Appliquez les recommandations générales suivantes pour la couleur :

- Utilisez la couleur avec parcimonie pour communiquer la hiérarchie et renforcer la marque.
- L’utilisation excessive d’une couleur d’accentuation unique appliquée aux éléments interactifs et non interactifs peut être source de confusion. Par exemple, évitez d’utiliser la même couleur pour les éléments sélectionnés et non sélectionnés dans un menu de navigation.
- Évitez les conflits inutiles avec des couleurs non Office.
- Utilisez vos propres couleurs de la marque pour créer une association avec votre service ou votre société.
- Assurez-vous que tout le texte est accessible. Assurez-vous qu’il existe un rapport de contraste 4,5:1 entre le texte de premier plan et l’arrière-plan.
- Gardez à l’esprit le daltonisme. Utilisez plusieurs couleurs pour indiquer l’interactivité et la hiérarchie.
- Consultez la rubrique [Guidelines Icons](../design/add-in-icons.md) pour en savoir plus sur la conception des icônes de commande de complément avec la couleur de l’icône Office palette couleurs.
