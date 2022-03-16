---
title: Ensembles de conditions requises pour le runtime partagé
description: Spécifie les plateformes et les applications Office qui la prise en charge des API SharedRuntime.
ms.date: 03/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 462a1248acadb35d4a4040fe9b997ede46844e4e
ms.sourcegitcommit: 856f057a8c9b937bfb37e7d81a6b71dbed4b8ff4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/16/2022
ms.locfileid: "63511281"
---
# <a name="shared-runtime-requirement-sets"></a>Ensembles de conditions requises pour le runtime partagé

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’ils nécessitent. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Les parties d’un Office qui exécutent du code JavaScript, telles que les volets Des tâches, les fichiers de fonctions lancés à partir de commandes de Excel et les fonctions personnalisées de Excel, peuvent partager un runtime JavaScript unique. Cela permet à tous les composants de partager un ensemble de variables globales, de partager un ensemble de bibliothèques chargées et de communiquer entre eux sans avoir à passer de messages via un stockage persistant. Pour plus d’informations, voir [Configurer votre Office pour utiliser un runtime JavaScript partagé](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).

Le tableau suivant répertorie l’ensemble de conditions requises SharedRuntime 1.1, les applications clientes Office qui la prise en charge, ainsi que les numéros de build ou de version de l’application Office.

| Ensemble de conditions requises | Office 2021 ou une Windows<br>(achat définitif) | Office pour Windows<br>(connecté à un abonnement Microsoft 365) | Office sur iPad<br>(connecté à un abonnement Microsoft 365) | Office sur Mac<br>(les deux abonnements<br> et achat Office sur Mac 2019 et ultérieur)  | Office sur le web | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| SharedRuntime 1.1  | Build 16.0.14326.20454 ou ultérieure | Version 2002 (build 12527.20092) ou version ultérieure | N/A | 16.35 ou version ultérieure | Février 2020 | N/A |

> [!IMPORTANT]
> Le runtime partagé JavaScript n’est à l’heure actuelle pas pris en charge on iPad ou les versions en achat définitif d’Office 2019 ou versions antérieures. Pour plus d’informations sur le support, consultez les sections suivantes.

## <a name="support-for-version-11-on-excel"></a>Prise en charge de la version 1.1 sur Excel

L’ensemble de conditions requises SharedRuntime 1.1 est publié pour Excel sur le Web, Windows et Mac.

## <a name="preview-support-for-version-11-on-word-and-powerpoint"></a>Prévisualisation de la prise en charge de la version 1.1 sur Word et PowerPoint

Le tableau suivant répertorie les builds d’application supplémentaires qui prisent en charge un aperçu du runtime JavaScript partagé. La version d’aperçu du runtime partagé est sujette à modification. Il n’est actuellement pas pris en charge pour une utilisation dans les environnements de production. Pour obtenir la dernière version, vous devez [rejoindre le programme Office Insider](https://insider.office.com/join). Un bon moyen de tester les fonctionnalités en préversion consiste à utiliser un abonnement Microsoft 365. Si vous n’avez pas déjà d’abonnement Microsoft 365, vous pouvez en obtenir un gratuitement en rejoignant le [Programme pour les développeurs Microsoft 365](https://developer.microsoft.com/office/dev-program).

|Application Office |Créer |
|-------------------|------|
|PowerPoint sur Windows |Build 16.0.13218.10000 ou ultérieure |
|PowerPoint sur Mac |Build 16.46.207.0 ou ultérieure |
|PowerPoint sur le web | Février 2022 |
|Word sur Windows |Build 16.0.13218.10000 ou ultérieure |
|Word sur Mac |Build 16.46.207.0 ou ultérieure |

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Voir aussi

- [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécifier les exigences en matière d’applications Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
