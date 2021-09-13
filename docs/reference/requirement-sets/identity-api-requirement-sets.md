---
title: Ensembles de conditions requises de l’API d’identité
description: Informations de l’ensemble de conditions requises de l’API d’identité Office les modules complémentaires.
ms.date: 01/26/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: d8a18ed8e7f78c5c83aeb2177a45c4fb46ba4a46
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152236"
---
# <a name="identity-api-requirement-sets"></a>Ensembles de conditions requises de l’API d’identité

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’ils nécessitent. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises de l’API d’identité, les applications clientes Office qui la prise en charge, ainsi que les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  | Office 2013 ou version ultérieure sous Windows<br>(achat définitif) | Office pour Windows<br>(connecté à un abonnement Microsoft 365) |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac<br>(connecté à un abonnement Microsoft 365)  | Office sur le web  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3  | N/A | 2008 (build 13127.20000) ou ultérieure | Bientôt disponible | 16.40 ou version ultérieure | Microsoft Office SharePoint Online et OneDrive\* |

\*Actuellement, l’ensemble de conditions requises est pris en charge Office sur le Web uniquement pour les documents ouverts à partir de Microsoft Office SharePoint Online et OneDrive.

> [!NOTE]
> Outlook : pour exiger l’ensemble d’API d’identité 1.3 dans le code de votre application, vérifiez s’il est pris en charge par `isSetSupported('IdentityAPI', '1.3')` l’appel. Sa déclaration dans le manifeste du Outlook n’est pas prise en charge. Vous pouvez également déterminer si l’API est prise en charge en vérifiant qu’elle n’est pas `undefined`. Pour plus d’informations, consultez [Utilisation des API d’un ensemble de conditions requises ultérieure](outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="identityapi-preview"></a>Aperçu IdentityAPI

Pour plus d’informations sur cette API, voir la version qui utilise Promises sur [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) ou la version qui utilise des rappels au [niveau de getAccessTokenAsync](/javascript/api/office/office.auth#getAccessTokenAsync_options__callback_).

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécifier les exigences en matière d’applications Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
