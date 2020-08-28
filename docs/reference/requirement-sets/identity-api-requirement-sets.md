---
title: Ensembles de conditions requises de l’API d’identité
description: Informations sur les conditions requises de l’API Identity pour les compléments Office.
ms.date: 07/30/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c2c6ea449cef08248a9ba79051b7c0c5f9baa600
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293540"
---
# <a name="identity-api-requirement-sets"></a>Ensembles de conditions requises de l’API d’identité

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API dont un complément a besoin. Pour plus d’informations, consultez la rubrique [versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises de l’API d’identité, les applications clientes Office qui prennent en charge l’ensemble de conditions requises, ainsi que les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  | Office 2013 ou version ultérieure sous Windows<br>(achat définitif) | Office pour Windows<br>(connecté à un abonnement Microsoft 365) |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac<br>(connecté à un abonnement Microsoft 365)  | Office sur le web  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Ensembles 1,3  | S/O | 2008 (Build 13127,20000) ou version ultérieure | Bientôt disponible | 16,40 ou version ultérieure | Août, 2020 * |

> \* Initialement, l’ensemble de conditions requises est pris en charge dans Office sur le Web uniquement pour les documents ouverts à partir de SharePoint Online et OneDrive.com. La prise en charge d’autres documents arrivera sur Office sur le Web plus tard dans 2020.

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="identityapi-preview"></a>Préversion ensembles

Pour plus d’informations sur cette API, consultez la version qui utilise les promesses sur [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) ou la version qui utilise les rappels sur [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécification des exigences en matière d’applications et d’API Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
