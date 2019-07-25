---
title: Ensembles de conditions requises de l’API JavaScript pour Word
description: Informations sur la configuration requise pour le complément Office sur les builds Word
ms.date: 07/17/2019
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: 4af7a9c14489d148ffdc06a68ad6c26bf326abc5
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/19/2019
ms.locfileid: "35804624"
---
# <a name="word-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour Word

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

## <a name="requirement-set-availability"></a>Disponibilité d’ensemble de conditions requises

Les compléments Word peuvent être exécutés dans différentes versions d’Office, notamment Office 2016 ou version ultérieure sur Windows, et Office sur le web, iPad et Mac. Le tableau suivant répertorie les ensembles de conditions requises pour Word, les applications Office hôte qui prennent en charge ces conditions et les numéros de build ou de version de ces applications.

> [!NOTE]
> Pour utiliser l’API dans un des jeux exigence numérotée, vous devez référencer la **production** de la bibliothèque sur le CDN : https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> Pour plus d’informations sur l’utilisation aperçu API, voir l’article[JavaScript d’Excel preview API](word-preview-apis.md).

|  Ensemble de conditions requises  |   Office pour Windows\*<br>(connecté à l’abonnement Office 365)  |  Office sur iPad<br>(connecté à l’abonnement Office 365)  |  Office sur Mac<br>(connecté à l’abonnement Office 365)  | Office sur le web  |
|:-----|-----|:-----|:-----|:-----|
| [Aperçu](word-preview-apis.md) | Veuillez utiliser la dernière version d’Office pour tester la préversion API (vous devrez peut-être rejoindre la [programme Office Insider](https://products.office.com/office-insider)) |
| [WordApi 1.3](word-api-1-3-requirement-set.md) | Version 1612 (Build 7668.1000) ou version ultérieure| Mars 2017, 2.22 ou version ultérieure | Mars 2017, 15.32 ou version ultérieure| Mars 2017 |
| [WordApi 1.2](word-api-1-2-requirement-set.md) | Mise à jour de décembre 2015, version 1601 (Build 6568.1000) ou version ultérieure | Janvier 2016, 1.18 ou version ultérieure | Janvier 2016, 15.19 ou version ultérieure| Septembre 2016 |
| [WordApi 1.1](word-api-1-1-requirement-set.md) | Version 1509 (Build 4266.1001) ou version ultérieure| Janvier 2016, 1.18 ou version ultérieure | Janvier 2016, 15.19 ou version ultérieure| Septembre 2016 |

> [!NOTE]
> Le numéro de build d’Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que l’ensemble d’exigences WordApi 1.1.

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour plus d’informations sur les versions et les numéros de build d’Office, voir :

- [Numéros de version et de build des canaux de réception des mises à jour pour les clients Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Où trouver le numéro de version et de build pour une application cliente Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Word](/javascript/api/word)
- [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Spécification des exigences en matière d’hôtes Office et d’API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifeste XML des compléments Office](/office/dev/add-ins/develop/add-in-manifests)
