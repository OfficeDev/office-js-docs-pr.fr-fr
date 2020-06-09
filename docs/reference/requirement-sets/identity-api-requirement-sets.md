---
title: Ensembles de conditions requises de l’API d’identité
description: Informations sur les conditions requises de l’API Identity pour les compléments Office.
ms.date: 04/16/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 1979f4ba27840885ac17ff21d81cbca3715c78bf
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611364"
---
# <a name="identity-api-requirement-sets"></a>Ensembles de conditions requises de l’API d’identité

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises de l’API de boîte de dialogue, les applications Office hôtes qui prennent en charge ces conditions et les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  | Office 2013 ou version ultérieure sous Windows<br>(achat définitif) | Office pour Windows<br>(connecté à l’abonnement Office 365) |  Office sur iPad<br>(connecté à l’abonnement Office 365)  |  Office sur Mac<br>(connecté à l’abonnement Office 365)  | Office sur le web  | SharePoint Online | OneDrive.com |Outlook.com et Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Préversion ensembles  | N/A | Préversion<b>*</b> | Bientôt disponible | Préversion<b>*</b> | Aperçu<b>* &#8224;</b> | Aperçu<b>* &#8224;</b>| Bientôt disponible | Bientôt disponible |

> **&#42;** Pendant la phase d’évaluation, l’API d’identité nécessite Office 365 (la version d’abonnement d’Office). Vous devez utiliser la version et le build mensuels les plus récents du canal du programme Insider. Vous devez participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://insider.office.com). Veuillez noter que lorsqu’un build passe au canal semi-annuel de production, la prise en charge des fonctionnalités d’aperçu, y compris l’authentification unique, est désactivée pour ce build.
>
> **&#8224;** Les compléments qui utilisent les API SSO sur ces plateformes ne fonctionnent que si l’administrateur client de l’utilisateur a accordé le consentement au complément. L’utilisateur ne peut pas accorder de consentement même à son propre profil Azure AD.

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
- [Spécification des exigences en matière d’hôtes Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
