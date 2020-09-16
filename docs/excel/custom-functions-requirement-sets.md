---
title: Ensembles de conditions requises pour les fonctions personnalisées
description: Détails sur les ensembles de conditions requises pour les fonctions personnalisées pour l’API JavaScript pour Excel.
ms.date: 09/14/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0860dd2d1b55376a85eadf04898d288d83b0205d
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819524"
---
# <a name="custom-functions-requirement-sets"></a>Ensembles de conditions requises pour les fonctions personnalisées

[Fonctions personnalisées](custom-functions-overview.md) utilisent des ensembles d’exigences distincts des API Excel JavaScript de base. Le tableau suivant répertorie les ensembles de conditions requises pour les fonctions personnalisées, les applications clientes Office prises en charge, ainsi que les versions ou le numéro de build de ces applications.

|  Ensemble de conditions requises  |  Office pour Windows<br>(connecté à un abonnement Microsoft 365)  |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac<br>(connecté à un abonnement Microsoft 365)  | Office sur le web |
|:-----|-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1,3 | 16.0.13127.20296 ou version ultérieure | Non pris en charge | 16.40.20081000 ou version ultérieure | Juillet 2020 |
| CustomFunctionsRuntime 1,2 | 16.0.12527.20194 ou version ultérieure | Non pris en charge | 16.34.20020900 ou version ultérieure | Janvier 2020 |
| CustomFunctionsRuntime 1.1 | 16.0.12527.20092 ou version ultérieure | Non pris en charge | 16,34 ou version ultérieure | Mai 2019 |

> [!NOTE]
> Les fonctions personnalisées d’Excel ne sont pas prises en charge dans Office 2019 ou version antérieure (achat unique).

## <a name="customfunctionsruntime-11-12-and-13"></a>CustomFunctionsRuntime 1,1, 1,2 et 1,3

Le CustomFunctionsRuntime 1,1 est la première version de l’API. L’ensemble de conditions requises 1,2 ajoute l' `CustomFunctions.Error` objet pour prendre en charge la gestion des erreurs. L’ensemble de conditions requises 1,3 ajoute la prise en charge de la [diffusion en continu XLL](make-custom-functions-compatible-with-xll-udf.md#custom-function-behavior-for-xll-compatible-functions) et de nouvelles `ErrorCode` options à l’objet [CustomFunctions. Error](/javascript/api/custom-functions-runtime/customfunctions.error) . 

## <a name="see-also"></a>Voir aussi

- [Documentation de référence sur les fonctions personnalisées](/javascript/api/custom-functions-runtime)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](../reference/requirement-sets/excel-api-requirement-sets.md)
