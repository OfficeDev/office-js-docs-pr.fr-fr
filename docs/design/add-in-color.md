---
title: Instructions concernant les couleurs pour les compléments Office
description: Découvrez comment utiliser des couleurs dans l’interface utilisateur d’un Office de l’interface utilisateur.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: a472dfd02787d68a5ce11a198d580aefe37ce315
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773222"
---
# <a name="color-guidelines-for-office-add-ins"></a>Instructions concernant les couleurs pour les compléments Office

La couleur est souvent utilisée pour mettre en évidence la marque et renforcer la hiérarchie visuelle. Elle permet d’identifier une interface et de guider les clients dans une expérience. Dans Office, la couleur est utilisée pour les mêmes objectifs mais elle est appliquée délibérément et au minimum. Elle ne surcharge jamais le contenu clients. Même lorsque chaque application Office est marquée avec sa propre couleur dominante, elle est utilisée avec parcimonie.

![Diagramme montrant le modèle de couleurs pour Office, Excel, Word et PowerPoint. Les couleurs principales Office sont le noir et le blanc, et les couleurs mineures sont gris clair, gris foncé et orange. La couleur dominante de Excel est vert, Word est bleu et PowerPoint orange.](../images/office-addins-color-schemes.png)

[Fabric Core inclut](fabric-core.md) un ensemble de couleurs de thème par défaut. Lorsque Fabric Core est appliqué à un Office dans les composants ou dans les dispositions, les mêmes objectifs s’appliquent. La couleur doit communiquer la hiérarchie, guidant ainsi les clients vers l’action sans interférer avec le contenu. Les couleurs de thème Fabric Core peuvent introduire une nouvelle couleur d’accentuer dans l’interface globale. Cette nouvelle accentuation peut entrer en conflit avec la personnalisation de l’application Office et interférer avec la hiérarchie. En d’autres termes, Fabric Core peut introduire une nouvelle couleur d’accentu utilisateur dans l’interface globale lorsqu’elle est utilisée dans un module. Cette nouvelle couleur de l’accentuation peut créer une confusion et interférer avec la hiérarchie globale. Envisagez des façons d’éviter les conflits et les interférences. Utilisez des accents neutres ou surécriture des couleurs de thème Fabric Core pour application Office la marque ou vos propres couleurs de marque.

Les applications Office permettent aux clients de personnaliser leurs interfaces en appliquant un thème de l’interface utilisateur d’Office. Les clients peuvent choisir entre quatre thèmes de l’interface utilisateur pour modifier le style des arrière-plans et des boutons dans Word, PowerPoint, Excel et les autres applications de la suite Office. Pour donner à vos add-ins l’impression d’être une partie naturelle de Office et répondre à la personnalisation, utilisez nos API de theming. Par exemple, les couleurs d’arrière-plan du volet des tâches deviennent gris foncé dans certains thèmes. Nos API de thèmes vous permettent de faire de même et d’ajuster le texte de premier plan pour garantir l’[accessibilité](../design/accessibility-guidelines.md).

> [!NOTE]
>
> - Pour les compléments de volet de tâches et de messagerie, utilisez la propriété [Context.officeTheme](/javascript/api/office/office.context) pour utiliser les thèmes correspondant à ceux des applications Office. Cette API est actuellement disponible dans Office 2016 ou ultérieure.
> - Pour plus d’informations sur les compléments de contenu pour PowerPoint, reportez-vous à l’article expliquant comment [utiliser des thèmes Office dans vos compléments PowerPoint](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md).

Appliquez les instructions générales suivantes pour la couleur.

- Utilisez la couleur avec parcimonie pour communiquer la hiérarchie et renforcer la marque.
- L’utilisation excessive d’une couleur d’accentuation unique appliquée aux éléments interactifs et non interactifs peut être source de confusion. Par exemple, évitez d’utiliser la même couleur pour les éléments sélectionnés et non sélectionnés dans un menu de navigation.
- Évitez les conflits inutiles avec des couleurs non Office.
- Utilisez vos propres couleurs de la marque pour créer une association avec votre service ou votre société.
- Assurez-vous que tout le texte est accessible. Assurez-vous qu’il existe un coefficient de contraste de 4,5:1 entre le texte au premier plan et l’arrière-plan.
- Pensez aux personnes atteintes de daltonisme : n’utilisez pas que des couleurs pour indiquer l’interactivité et la hiérarchie.
- Reportez-vous aux [instructions](../design/add-in-icons.md) relatives aux icônes pour en savoir plus sur la conception d’icônes de commande de Office palette de couleurs d’icône.
