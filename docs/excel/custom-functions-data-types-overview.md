---
title: Présentation des fonctions personnalisées et des types de données
description: Utilisez des types de données Excel avec vos fonctions personnalisées et compléments Office.
ms.date: 11/03/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 91d2fb21aae57ed7a5777136f3c4540925f339c8
ms.sourcegitcommit: 210251da940964b9eb28f1071977ea1fe80271b4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/05/2021
ms.locfileid: "60793574"
---
# <a name="use-data-types-with-custom-functions-in-excel-preview"></a>Utiliser des types de données avec des fonctions personnalisées dans Excel (préversion)

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

Les types de données étendent l'API JavaScript Excel pour prendre en charge les types de données au-delà des quatre types de données d'origine (chaîne, nombre, booléen et erreur). Les types de données incluent la prise en charge des images Web, des valeurs numériques formatées, des valeurs d'entité et des tableaux au sein des valeurs d'entité.

Ces types de données amplifient la puissance des fonctions personnalisées, car les fonctions personnalisées acceptent les types de données comme valeurs d'entrée et de sortie. Vous pouvez générer des types de données via des fonctions personnalisées ou utiliser des types de données existants comme arguments de fonction dans les calculs. Une fois le schéma JSON d'un type de données défini, ce schéma est conservé tout au long des calculs de fonction personnalisée.

Pour en savoir plus sur l'utilisation des types de données avec un complément Excel, consultez [Présentation des types de données dans les compléments Excel](excel-data-types-overview.md). Pour en savoir plus sur l'intégration des types de données personnalisés à vos fonctions personnalisées, consultez [Fonctions personnalisées et concepts de base des types de données](custom-functions-data-types-concepts.md).

## <a name="see-also"></a>Voir aussi

* [Présentation des types de données dans les compléments Excel](excel-data-types-overview.md)
* [Concepts de base des types de données Excel](excel-data-types-concepts.md)
* [Concepts de base des fonctions personnalisées et des types de données](custom-functions-data-types-concepts.md)
* [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
