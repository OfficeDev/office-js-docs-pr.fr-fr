---
title: Instructions concernant la typographie pour les compléments Office
description: ''
ms.date: 06/27/2018
localization_priority: Normal
ms.openlocfilehash: 5e43d538f87da3a318ce24e14f48acd8f7e15b88
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32446947"
---
# <a name="typography"></a>Typographie

Segoe est la police standard pour Office. Utilisez-la dans votre complément pour être en adéquation avec les volets des tâches, les boîtes de dialogue et les objets de contenu d’Office. Office UI Fabric vous donne accès à Segoe. Il fournit un dégradé de polices complet de Segoe avec de nombreuses variations, d’épaisseur de police et de taille, dans des classes CSS pratiques. Toutes les tailles et épaisseurs de police d’Office UI Fabric n’ont pas une belle apparence dans un complément Office. Pour une intégration harmonieuse ou pour éviter les incompatibilités, envisagez d’utiliser un sous-ensemble du dégradé de polices de Fabric. Voici une liste des classes de base de la structure que nous vous recommandons d’utiliser dans les compléments Office.

|Exemple |Classe |Taille |Pondération |Utilisation recommandée |
|------ |----- |---- |------ |----------------- |
|![Image de texte Hero](../images/add-in-typeramp-hero.png)|.ms-font-xxl |28 px | Segoe Light |<ul><li>Cette classe est plus grande que tous les autres éléments typographiques dans Office. Utilisez-la avec parcimonie pour éviter une hiérarchie visuelle non valide.</li><li>Évitez d’utiliser de longues chaînes dans des espaces limités.</li><li>Laissez suffisamment d’espaces blancs autour du texte en utilisant cette classe.</li><li>Couramment utilisée pour les premiers messages, éléments hero ou autres appels à l’action.</li></ul> |
|![Image de texte Hero](../images/add-in-typeramp-title.png)|.ms-font-xl |21 px |Segoe Light | <ul><li>Cette classe correspond au titre du volet des tâches des applications Office.</li><li>Utilisez-la avec parcimonie pour éviter une hiérarchie typographique plate.</li><li>Couramment utilisée comme élément de niveau supérieur (titres de contenu, de page ou de boîte de dialogue).</li></ul> |
|![Image de texte Hero](../images/add-in-typeramp-subtitle.png)|.ms-font-l |17 px |Segoe Semilight | <ul><li>Cette classe est le premier point en dessous des titres.</li><li>Couramment utilisée comme sous-titre, élément de navigation ou en-tête de groupe.</li><ul> |
|![Image de texte Hero](../images/add-in-typeramp-body.png)|.ms-font-m |14 px |Segoe Regular |<ul><li>Couramment utilisée comme corps de texte dans les compléments.</li><ul>|
|![Image de texte Hero](../images/add-in-typeramp-caption.png)|.ms-font-xs |11 px | Segoe Regular |<ul><li>Couramment utilisée pour le texte secondaire ou tertiaire (horodatages, signatures, légendes ou étiquettes de champ).</li><ul>|
|![Image de texte Hero](../images/add-in-typeramp-annotation.png)|.ms-font-mi |10 px |Segoe Semibold |<ul><li>Le plus petit niveau dans le dégradé de polices doit être rarement utilisé. Il est disponible lorsque la lisibilité n’est pas requise.</li><ul>|

> [!NOTE]
> La couleur du texte n’est pas incluse dans ces classes de base. Utilisez le « neutre primaire » de Fabric pour la plupart du texte affichés sur des arrière-plans blancs.
