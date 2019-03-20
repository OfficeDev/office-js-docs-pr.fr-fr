---
title: Ensembles de conditions requises de l’API d’identité
description: ''
ms.date: 03/19/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: e1f646df7756f16467d1191f926f831ba375752e
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691159"
---
# <a name="identity-api-requirement-sets"></a>Ensembles de conditions requises de l’API d’identité

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises de l’API de boîte de dialogue, les applications Office hôtes qui prennent en charge ces conditions et les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  | Office 2013 ou version ultérieure pour Windows | Office 365 pour Windows   |  Office 365 pour iPad  |  Office 365 pour Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com et Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | N/A **& # 42;** | Aperçu **&#42;** | Bientôt disponible | Aperçu **&#42;**| Aperçu | Aperçu| Bientôt disponible | Bientôt disponible |

> **& # 42;** Pendant la phase d'évaluation, l'API d'identité nécessite Office 365 (la version d'abonnement d'Office). Vous devez utiliser la version et le build mensuels les plus récents du canal du programme Insider. Vous devez participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://products.office.com/office-insider?tab=tab-1). Veuillez noter que lorsqu’un build passe au canal semi-annuel de production, la prise en charge des fonctionnalités d’aperçu, y compris l’authentification unique, est désactivée pour ce build.

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

- [Numéros de version et de build des canaux de réception des mises à jour pour les clients Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Où trouver le numéro de version et de build pour une application cliente Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des APIs communes, voir[Ensembles de conditions requises des APIs communes pour Office](office-add-in-requirement-sets.md).

## <a name="identityapi-11"></a>IdentityAPI 1.1

IdentityAPI 1.1 à connexion unique est la première version de l’API. Pour plus d’informations sur cette API, voir la[référence de l’API de l’authentification unique SSO](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) section d’[Activation de l’authentification unique dans un complément](/office/dev/add-ins/develop/sso-in-office-add-ins).

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Spécification des exigences en matière d’hôtes Office et d’API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifeste XML des compléments Office](/office/dev/add-ins/develop/add-in-manifests)
