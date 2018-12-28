---
title: Boîtes de dialogue dans les compléments Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 5874e93c79d019adbd7ca5827a5b5e22f742190a
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457669"
---
# <a name="dialog-boxes-in-office-add-ins"></a>Boîtes de dialogue dans les compléments Office
 
Les boîtes de dialogue sont des surfaces qui flottent au-dessus de la fenêtre active de l’application Office. Vous pouvez utiliser les boîtes de dialogue afin de fournir un espace supplémentaire sur l’écran pour les tâches comme les pages de connexion impossibles à ouvrir directement dans un volet des tâches, ou pour les demandes de confirmation d’une action effectuée par un utilisateur, ou pour afficher des vidéos qui peuvent être trop petites si confinées à un volet des tâches.

*Figure 1. Mise en page type pour une boîte de dialogue*

![Exemple d’image affichant une mise en page par défaut pour une boîte de dialogue](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a>Meilleures pratiques

|**À faire**|**À ne pas faire**|
|:-----|:--------|
|<ul><li>Inclure un titre descriptif qui inclut le nom de votre complément, ainsi que la tâche en cours.</li></ul>|<ul><li>Ne pas ajouter le nom de votre société au titre.</li></ul>|
||<ul><li>Ne pas ouvrir une boîte de dialogue, sauf si le scénario l’exige.</li></ul>|

## <a name="implementation"></a>Implémentation

Pour voir un exemple relatif à l’implémentation d’une boîte de dialogue, consultez [Exemple d’API de boîte de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) dans GitHub.

## <a name="see-also"></a>Voir aussi

- [Ressources de développement GitHub](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Objet Dialogue](https://docs.microsoft.com/javascript/api/office/office.dialog)


