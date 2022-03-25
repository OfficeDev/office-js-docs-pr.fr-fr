---
title: Ensembles de conditions requises de l’API JavaScript pour Word
description: Informations sur l’ensemble des exigences du complément Office pour Word.
ms.date: 01/14/2022
ms.prod: word
ms.localizationpriority: high
ms.openlocfilehash: 25a698a82669210a596026807d9e30be38a68762
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746175"
---
# <a name="word-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour Word

Les ensembles de conditions requises sont des groupes nommés des membres de l’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si une application Office prend en charge les API requises par un complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

## <a name="requirement-set-availability"></a>Disponibilité d’ensemble de conditions requises

Les compléments Word peuvent être exécutés dans différentes versions d’Office, notamment Office 2016 ou version ultérieure sur Windows, et Office sur le web, iPad et Mac. Le tableau suivant répertorie les ensembles de conditions requises pour Word, les applications clientes Office qui prennent en charge l’ensemble des conditions requises et les numéros de création ou de version de ces applications.

> [!NOTE]
> Pour utiliser des API dans l’un des ensembles de conditions requises numérotées, vous devez référencer la bibliothèque de **production** sur le [réseau de distribution de contenu (CDN) Office.js](https://appsforoffice.microsoft.com/lib/1/hosted/office.js).
>
> Pour plus d’informations sur l’utilisation aperçu API, voir l’article[JavaScript de Word preview API](word-preview-apis.md).

|  Ensemble de conditions requises  |   Office pour Windows\*<br>(connecté à un abonnement Microsoft 365)  |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac<br>(connecté à un abonnement Microsoft 365)  | Office sur le web  |
|:-----|-----|:-----|:-----|:-----|
| [Aperçu](word-preview-apis.md) | Veuillez utiliser la dernière version d’Office pour tester la préversion API (vous devrez peut-être rejoindre la [programme Office Insider](https://insider.office.com)) |
| [WordApi 1.3](word-api-1-3-requirement-set.md) | Version 1612 (Build 7668.1000) ou version ultérieure| Mars 2017, 2.22 ou version ultérieure | Mars 2017, 15.32 ou version ultérieure| Mars 2017 |
| [WordApi 1.2](word-api-1-2-requirement-set.md) | Mise à jour de décembre 2015, version 1601 (Build 6568.1000) ou version ultérieure | Janvier 2016, 1.18 ou version ultérieure | Janvier 2016, 15.19 ou version ultérieure| Septembre 2016 |
| [WordApi 1.1](word-api-1-1-requirement-set.md) | Version 1509 (Build 4266.1001) ou version ultérieure| Janvier 2016, 1.18 ou version ultérieure | Janvier 2016, 15.19 ou version ultérieure| Septembre 2016 |

> [!NOTE]
> Les versions sans abonnement d'Office prennent en charge l'ensemble des conditions requises suivantes :
>
> - Office 2019 et Office 2021 prennent en charge WordApi 1.3 et versions antérieures.
> - Office 2016 prend uniquement en charge l'ensemble des conditions requises de WordApi 1.1.

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour plus d’informations sur les versions et les numéros de build d’Office, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="see-also"></a>Articles associés

- [Documentation référence de l’API JavaScript pour Word](/javascript/api/word)
- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécifier les exigences en matière d’applications Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
