---
title: Séries de conditions requises pour ouvrir une fenêtre de navigateur
description: Spécifie les plateformes Office et les builds qui ouvrent la prise en charge de l’API openBrowserWindow.
ms.date: 10/05/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 65088c6a8fa7c7f9bd3e53da0d2b1755b7ec22e6
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138785"
---
# <a name="open-browser-window-api-requirement-sets"></a>Ouvrir les ensembles de conditions requises de l’API Fenêtre du navigateur

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

L’ensemble d’API OpenBrowserWindow permet aux applications d’ouvrir un navigateur pour accomplir des tâches qui ne peuvent pas toujours être réalisées dans le contrôle webview en bac à sable (sandbox) au sein du module lui-même. par exemple, le téléchargement d’un fichier PDF lorsque le contrôle webview est fourni par Microsoft Edge.

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises de l’API OpenBrowserWindow, les applications hôtes Office qui la prise en charge, ainsi que les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  | Office 2021 ou une Windows<br>(achat définitif) | Office pour Windows<br>(connecté à l’abonnement Microsoft 365) |  Office sur iPad<br>(connecté à l’abonnement Microsoft 365)  |  Office sur Mac<br>(connecté à l’abonnement Microsoft 365)  | Office sur le web  |  Office Online Server  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| OpenBrowserWindowApi 1.1  | Build 16.0.14326.20454 ou ultérieure | Version 1810 (build 16.0.11001.20074) ou version ultérieure | 16.0.0.0 ou ultérieur | 16.0.0.0 ou ultérieur | N/A | N/A|

> [!NOTE]
> L’ensemble de conditions requises OpenBrowserWindowApi est disponible uniquement comme suit :
>
> - Excel, PowerPoint, Word : Windows, Mac, iPad
> - Outlook : Windows, Mac

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

- [Numéros de version et de build des versions de canal de mise à jour Microsoft 365 Apps](/officeupdates/update-history-microsoft365-apps-by-date)
- [Quelle est la version d’Office que j’utilise ?](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Où trouver la version et le numéro de build d’une application Office client](/officeupdates/update-history-microsoft365-apps-by-date)
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="openbrowserwindowapi-11"></a>OpenBrowserWindowApi 1.1

OpenBrowserWindowApi 1.1 est la première version de l’API. Pour plus d’informations sur l’API, voir [Office.context.ui](/javascript/api/office/office.context#ui) reference topic.

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécification des exigences en matière d’hôtes Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
