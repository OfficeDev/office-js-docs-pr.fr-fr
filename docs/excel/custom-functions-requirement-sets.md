---
title: Ensembles de conditions requises fonctions personnalisées
description: Détails sur les ensembles de conditions requises fonctions personnalisées pour Excel API JavaScript.
ms.date: 09/14/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5fbb280176b6c222bedb7cefe14c3caa92148032b59cc1d6c0942dde1f52a3aa
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57079470"
---
# <a name="custom-functions-requirement-sets"></a>Ensembles de conditions requises fonctions personnalisées

[Fonctions personnalisées](custom-functions-overview.md) utilisent des ensembles d’exigences distincts des API Excel JavaScript de base. Le tableau suivant répertorie les ensembles de conditions requises des fonctions personnalisées, les applications clientes Office prise en charge et les versions ou le numéro de build de ces applications.

|  Ensemble de conditions requises  |  Office pour Windows<br>(connecté à un abonnement Microsoft 365)  |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac<br>(connecté à un abonnement Microsoft 365)  | Office sur le web |
|:-----|-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1.3 | 16.0.13127.20296 ou ultérieur | Non pris en charge | 16.40.20081000 ou ultérieure | Juillet 2020 |
| CustomFunctionsRuntime 1.2 | 16.0.12527.20194 ou ultérieur | Non pris en charge | 16.34.20020900 ou ultérieure | Janvier 2020 |
| CustomFunctionsRuntime 1.1 | 16.0.12527.20092 ou ultérieure | Non pris en charge | 16.34 ou ultérieure | Mai 2019 |

> [!NOTE]
> Excel fonctions personnalisées ne sont pas Office 2019 ou antérieures (achat à usage seul).

## <a name="customfunctionsruntime-11-12-and-13"></a>CustomFunctionsRuntime 1.1, 1.2 et 1.3

CustomFunctionsRuntime 1.1 est la première version de l’API. L’ensemble de conditions requises 1.2 ajoute `CustomFunctions.Error` l’objet pour prendre en charge la gestion des erreurs. L’ensemble de conditions requises 1.3 ajoute la prise en charge de la diffusion en continu [XLL](make-custom-functions-compatible-with-xll-udf.md#custom-function-behavior-for-xll-compatible-functions) et de nouvelles options à `ErrorCode` [l’objet CustomFunctions.Error.](/javascript/api/custom-functions-runtime/customfunctions.error) 

## <a name="see-also"></a>Voir aussi

- [Documentation de référence sur les fonctions personnalisées](/javascript/api/custom-functions-runtime)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](../reference/requirement-sets/excel-api-requirement-sets.md)
