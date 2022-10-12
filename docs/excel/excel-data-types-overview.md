---
title: Vue d’ensemble des types de données dans les compléments Excel
description: Les types de données dans l’API JavaScript Excel permettent aux développeurs de compléments Office d’utiliser des valeurs numériques mises en forme, des images web, des entités, des tableaux au sein d’entités et des erreurs améliorées en tant que types de données.
ms.date: 10/10/2022
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 2d19eacc23d64f472f32363fc93155b6e023ba04
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2022
ms.locfileid: "68540976"
---
# <a name="overview-of-data-types-in-excel-add-ins"></a>Vue d’ensemble des types de données dans les compléments Excel

Les types de données organisent des structures de données complexes en tant qu’objets. Cela inclut les valeurs numériques mises en forme, les images web et les entités en tant que [cartes d’entité](excel-data-types-entity-card.md).

Avant l’ajout des types de données, l’API JavaScript Excel prenait en charge les types de données chaîne, nombre, booléen et d’erreur. La couche de mise en forme de l’interface utilisateur Excel est capable d’ajouter des devises, des dates et d’autres types de mise en forme aux cellules qui contiennent les quatre types de données d’origine, mais cette couche de mise en forme contrôle uniquement l’affichage des types de données d’origine dans l’interface utilisateur Excel. La valeur du nombre sous-jacent n’est pas modifiée, même lorsqu’une cellule de l’interface utilisateur Excel est mise en forme en tant que devise ou date. Cet écart entre une valeur sous-jacente et l’affichage mis en forme dans l’interface utilisateur Excel peut se traduire par une confusion et des erreurs pendant les calculs du complément. Les API de types de données constituent une solution à cet écart.

Les types de données étendent la prise en charge de l’API JavaScript Excel au-delà des quatre types de données d’origine (chaîne, nombre, booléen et erreur) pour inclure des [images web](excel-data-types-concepts.md#web-image-values), [des valeurs numériques mises en forme](excel-data-types-concepts.md#formatted-number-values), [des entités](excel-data-types-concepts.md#entity-values), des tableaux dans des entités et des [types de données d’erreur](excel-data-types-concepts.md#improved-error-support) améliorés en tant que structures de données flexibles. Ces types, qui permettent de nombreuses expériences de [types de données liées](https://support.microsoft.com/office/what-linked-data-types-are-available-in-excel-6510ab58-52f6-4368-ba0f-6a76c0190772), offrent une précision et une simplicité lors des calculs du complément et étendent le potentiel des compléments Excel au-delà d’une grille à 2 dimensions.

Pour savoir comment utiliser les API de types de données, commencez par l’article sur les [concepts de base des types de données Excel](excel-data-types-concepts.md) .

## <a name="data-types-and-custom-functions"></a>Types de données et fonctions personnalisées

Les types de données améliorent la puissance des fonctions personnalisées. Les fonctions personnalisées acceptent les types de données comme entrées et sorties de fonctions personnalisées et les fonctions personnalisées utilisent le même schéma JSON pour les types de données que l’API JavaScript Excel. Ce schéma JSON de types de données est conservé à mesure que les fonctions personnalisées calculent et évaluent. Pour en savoir plus sur l’intégration des types de données à vos fonctions personnalisées, consultez[Fonctions personnalisées et types de données.](custom-functions-data-types-concepts.md)

## <a name="see-also"></a>Voir aussi

- [Concepts de base des types de données Excel](excel-data-types-concepts.md)
- [Utiliser des cartes avec des types de données de valeur d’entité](excel-data-types-entity-card.md)
- [Référence de l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Fonctions personnalisées et types de données](custom-functions-data-types-concepts.md)