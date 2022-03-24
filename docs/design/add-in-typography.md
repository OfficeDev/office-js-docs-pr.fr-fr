---
title: Instructions concernant la typographie pour les compléments Office
description: Découvrez les polices et les tailles de police à utiliser dans Office de police.
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: f63d4b6816b916dc52711a8f4b11e826efd58105
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63742960"
---
# <a name="typography"></a>Typographie

Segoe est la police standard pour Office. Utilisez-la dans votre complément pour être en adéquation avec les volets des tâches, les boîtes de dialogue et les objets de contenu d’Office. [Fabric Core vous](fabric-core.md) donne accès à Segoe. Il fournit un dégradé de polices complet de Segoe avec de nombreuses variations, d’épaisseur de police et de taille, dans des classes CSS pratiques. Toutes les tailles et poids Fabric Core n’auront pas une apparence excellente dans un Office de base. Pour s’ajuster harmonieusement ou éviter les conflits, envisagez d’utiliser un sous-ensemble de la ramp de type Fabric Core. Le tableau suivant répertorie les classes de base de Fabric Core que nous vous recommandons d’utiliser dans Office de base.

> [!NOTE]
> La couleur du texte n’est pas incluse dans ces classes de base. Utilisez le « principal neutre » de Fabric Core pour la plupart du texte sur des arrière-plans blancs.
>
> Pour en savoir plus sur la typographie disponible, voir [Web Typography](https://developer.microsoft.com/fluentui#/styles/web/typography).

|Type |Classe |Taille |Pondération |Utilisation recommandée |
|------ |----- |---- |------ |----------------- |
|Bannière|.ms-font-xxl |28 px | Segoe Light |<ul><li>Cette classe est plus grande que tous les autres éléments typographiques dans Office. Utilisez-la avec parcimonie pour éviter une hiérarchie visuelle non valide.</li><li>Évitez d’utiliser de longues chaînes dans des espaces limités.</li><li>Laissez suffisamment d’espaces blancs autour du texte en utilisant cette classe.</li><li>Couramment utilisée pour les premiers messages, éléments hero ou autres appels à l’action.</li></ul> |
|Title|.ms-font-xl |21 px |Segoe Light | <ul><li>Cette classe correspond au titre du volet des tâches des applications Office.</li><li>Utilisez-la avec parcimonie pour éviter une hiérarchie typographique plate.</li><li>Couramment utilisée comme élément de niveau supérieur (titres de contenu, de page ou de boîte de dialogue).</li></ul> |
|Subtitle|.ms-font-l |17 px |Segoe Semilight | <ul><li>Cette classe est le premier point en dessous des titres.</li><li>Couramment utilisée comme sous-titre, élément de navigation ou en-tête de groupe.</li><ul> |
|Corps|.ms-font-m |14 px |Segoe Regular |<ul><li>Couramment utilisée comme corps de texte dans les compléments.</li><ul>|
|Légende|.ms-font-xs |11 px | Segoe Regular |<ul><li>Couramment utilisée pour le texte secondaire ou tertiaire (horodatages, signatures, légendes ou étiquettes de champ).</li><ul>|
|Annotation|.ms-font-mi |10 px |Segoe Semibold |<ul><li>Le plus petit niveau dans le dégradé de polices doit être rarement utilisé. Il est disponible lorsque la lisibilité n’est pas requise.</li><ul>|
