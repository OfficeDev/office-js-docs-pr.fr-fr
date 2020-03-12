---
title: Ensembles de conditions requises pour l’exécution partagée
description: Spécifie les plateformes et les hôtes Office qui prennent en charge les API SharedRuntime.
ms.date: 03/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 9750dd2e20a9f2426c7faf3864a2fccac5c11a80
ms.sourcegitcommit: 05b73cdec5f4db7f0b8d48a5a552ee296a0332ca
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42600675"
---
# <a name="shared-runtime-requirement-sets"></a>Ensembles de conditions requises pour l’exécution partagée

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Les parties d’un complément Office qui exécutent du code JavaScript, telles que des volets de tâches, des fichiers de fonctions lancés à partir de commandes de complément et des fonctions personnalisées Excel, peuvent partager un seul Runtime JavaScript. Cela permet à toutes les parties de partager un ensemble de variables globales, de partager un ensemble de bibliothèques chargées et de communiquer les uns avec les autres sans avoir à transmettre de messages via un stockage persistant.

Le tableau suivant répertorie l’ensemble de conditions requises SharedRuntime 1,1, les applications hôtes Office qui prennent en charge cet ensemble de conditions requises, ainsi que les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  |  Office 2013 (ou version ultérieure) sur Windows<br>(achat définitif) | Office pour Windows<br>(connecté à l’abonnement Office 365)   |  Office sur iPad<br>(connecté à l’abonnement Office 365)  |  Office sur Mac<br>(connecté à l’abonnement Office 365)  | Office sur le web  | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| SharedRuntime 1,1  | S/O | Version 2002 (Build 12527,20092) ou version ultérieure | S/O | 16,35 ou version ultérieure | Février 2020 | S/O |

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécification des exigences en matière d’hôtes Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
