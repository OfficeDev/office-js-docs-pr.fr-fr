---
title: Ensembles de conditions requises fonctions personnalisées
description: Détails sur les ensembles de conditions requises fonctions personnalisées pour Excel API JavaScript.
ms.date: 02/15/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 7558035b6b151977e985ec04ed1fa84c116f0886
ms.sourcegitcommit: 789545a81bd61ec2e7adef2bc24c06b5be113b00
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/18/2022
ms.locfileid: "62892516"
---
# <a name="custom-functions-requirement-sets"></a>Ensembles de conditions requises fonctions personnalisées

[Fonctions personnalisées](../../excel/custom-functions-overview.md) utilisent des ensembles d’exigences distincts des API Excel JavaScript de base. Le tableau suivant répertorie les ensembles de conditions requises des fonctions personnalisées, les applications clientes Office prise en charge et les versions ou le numéro de build de ces applications.

|  Ensemble de conditions requises  |  Office 2021 ou une Windows<br>(achat définitif)  |  Office pour Windows<br>(connecté à un abonnement Microsoft 365)  |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac<br>(les deux abonnements<br> et achat Office sur Mac 2021 et ultérieur)  | Office sur le web |
|:-----|:-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1.3 | 16.0.14326.20454 ou ultérieur | 16.0.13127.20296 ou ultérieur | Non pris en charge | 16.40.20081000 ou ultérieure | Juillet 2020 |
| CustomFunctionsRuntime 1.2 | 16.0.14326.20454 ou ultérieur | 16.0.12527.20194 ou ultérieur | Non pris en charge | 16.34.20020900 ou ultérieure | Janvier 2020 |
| CustomFunctionsRuntime 1.1 | 16.0.14326.20454 ou ultérieur | 16.0.12527.20092 ou ultérieure | Non pris en charge | 16.34 ou ultérieure | Mai 2019 |

## <a name="customfunctionsruntime-11-12-and-13"></a>CustomFunctionsRuntime 1.1, 1.2 et 1.3

CustomFunctionsRuntime 1.1 est la première version de l’API. L’ensemble de conditions requises 1.2 ajoute l’objet `CustomFunctions.Error` pour prendre en charge la gestion des erreurs. L’ensemble de conditions requises 1.3 ajoute la prise en charge de [la diffusion](../../excel/make-custom-functions-compatible-with-xll-udf.md#custom-function-behavior-for-xll-compatible-functions) `ErrorCode` en continu XLL et de nouvelles options à l’objet [CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) .

## <a name="see-also"></a>Voir aussi

- [Documentation de référence sur les fonctions personnalisées](/javascript/api/custom-functions-runtime)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
