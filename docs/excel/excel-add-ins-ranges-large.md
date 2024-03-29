---
title: Lire ou écrire dans de grandes plages à l’aide de l Excel API JavaScript
description: Découvrez comment lire ou écrire dans de grandes plages à l’Excel’API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 64b47c59e231b0ef40f81d670c511eb7836bd204
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745306"
---
# <a name="read-or-write-to-a-large-range-using-the-excel-javascript-api"></a>Lire ou écrire dans une grande plage à l’aide de l Excel API JavaScript

Cet article explique comment gérer la lecture et l’écriture dans de grandes plages à l’Excel’API JavaScript.

## <a name="run-separate-read-or-write-operations-for-large-ranges"></a>Exécuter des opérations de lecture ou d’écriture distinctes pour des plages de grande taille

Si une plage contient un grand nombre de cellules, valeurs, formats numériques ou formules, il est possible qu’il ne soit pas possible d’exécuter des opérations d’API sur cette plage. L’API essaie toujours d’exécuter au mieux l’opération demandée sur une plage (par exemple, pour extraire ou écrire des données spécifiées), mais essayer d’effectuer des opérations de lecture ou d’écriture pour une grande plage peut provoquer une erreur d’API en raison de l’utilisation des ressources excessive. Pour éviter ces erreurs, nous vous recommandons d’exécuter des opérations de lecture ou d’écriture distinctes pour des sous-ensembles plus petits d’une grande plage, au lieu d’essayer d’exécuter une seule opération de lecture ou d’écriture sur une grande plage.

Pour plus d’informations sur les limitations du système, voir la section « Excel » des limites de ressources et l’optimisation des performances pour les Office [de gestion](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).

### <a name="conditional-formatting-of-ranges"></a>Mise en forme conditionnelle de plages

Des plages peuvent présenter une mise en forme de cellules individuelles en fonction de certaines conditions. Pour plus d’informations à ce sujet, consultez l’article [Appliquer une mise en forme conditionnelle à des plages Excel](excel-add-ins-conditional-formatting.md).

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide Excel API JavaScript](excel-add-ins-cells.md)
- [Lire ou écrire dans une plage non limite à l’aide de l Excel API JavaScript](excel-add-ins-ranges-unbounded.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
