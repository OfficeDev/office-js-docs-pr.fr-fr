---
title: Ensembles de conditions requises d’origine de boîte de dialogue
description: En savoir plus sur les ensembles de conditions requises d’origine de la boîte de dialogue.
ms.date: 07/22/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 24513823eb60435359d5d7307a11a192fece2015
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53774321"
---
# <a name="dialog-origin-requirement-sets"></a>Ensembles de conditions requises d’origine de boîte de dialogue

Les ensembles de conditions requises sont des groupes nommés de membres de l’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’un complément nécessite. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises d’origine de la boîte de dialogue, les applications clientes Office qui la prise en charge, ainsi que les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  | Office 2013 sur Windows<br>(achat définitif) | Office 2016 sur Windows<br>(achat définitif) | Office 2019 ou une Windows<br>(achat définitif) | Office pour Windows<br>(abonnement) |  Office sur iPad<br>(abonnement)  |  Office sur Mac<br>(abonnement)  | Office sur le web  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogOrigin 1.1  | Créer<br>15.0.5371.1000<br>ou ultérieure | Créer<br>16.0.5200.1000<br>ou ultérieure | Créer<br>À déterminer<br>ou ultérieure | À déterminer | 2.52 ou ultérieure | 16.52 ou ultérieure | Juillet 2021 | Version 2108<br>(Build 10377.1000)<br>ou ultérieure |

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="dialog-origin-11"></a>Dialog Origin 1.1

Dialog Origin 1.1 est la première version de l’API. Il assure la prise en charge de la messagerie entre domaines entre une boîte de dialogue et sa page parente. Pour plus d’informations sur ces API, voir [la rubrique Office.ui.](/javascript/api/office/office.ui)

## <a name="see-also"></a>Voir aussi

- [Utiliser l’API de boîte de dialogue Office dans les compléments Office](../../develop/dialog-api-in-office-add-ins.md)
- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécifier les exigences en matière d’applications Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
