---
title: Ensembles de conditions requises de l’API d’identité
description: Informations de l’ensemble de conditions requises de l’API d’identité Office les modules complémentaires.
ms.date: 02/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: bff7d75d538922f6d5d5d05a01306a4ba2ec836c
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744925"
---
# <a name="identity-api-requirement-sets"></a>Ensembles de conditions requises de l’API d’identité

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’ils nécessitent. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises de l’API d’identité, les applications clientes Office qui la prise en charge, ainsi que les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  | Office 2021 ou une Windows<br>(achat définitif) | Office pour Windows<br>(connecté à un abonnement Microsoft 365) |  Office sur iPad<br>(connecté à un abonnement Microsoft 365)  |  Office sur Mac<br>(les deux abonnements<br> et achat Office sur Mac 2019 et ultérieur)   | Office sur le web  |
|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3  | Build 16.0.14326.20454 ou ultérieure | Version 2008 (build 13127.20000) ou version ultérieure | Non pris en charge | 16.40 ou version ultérieure | Microsoft Office SharePoint Online et OneDrive\* |

\*Actuellement, l’ensemble de conditions requises est pris en charge Office sur le Web uniquement pour les documents ouverts à partir de Microsoft Office SharePoint Online et OneDrive.

## <a name="outlook-and-identity-api-requirement-sets"></a>Outlook et ensembles de conditions requises de l’API d’identité

[!INCLUDE [How to use the Identity 1.3 requirement set in Outlook add-ins](../../includes/outlook-identity-13-note.md)]

> [!NOTE]
> Dans un complément Outlook utilisant l’activation basée sur des événements, [l’interface OfficeRuntime.Auth](/javascript/api/office-runtime/officeruntime.auth) est prise en charge sur Office sur Windows version 2108 (build 14326.20258) ou version ultérieure. Le [Office. L’interface](/javascript/api/office/office.auth) d’th est prise en charge sur la version 2109 (build 14425.10000) ou version ultérieure. Pour plus d’informations en fonction de votre version, consultez la page historique des mises à jour pour [Office 2021](/officeupdates/update-history-office-2021) ou [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) et comment trouver votre [version du client Office](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19) et le canal de mise à jour.

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécifier les exigences en matière d’applications Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
