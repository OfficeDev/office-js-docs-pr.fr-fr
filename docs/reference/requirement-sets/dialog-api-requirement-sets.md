---
title: Ensembles de conditions requises de l’API de dialogue
description: En savoir plus sur les ensembles de conditions requises de l’API Dialog.
ms.date: 09/14/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 79b6960387519ac3c8b41b0b31cf6f40b5e7e067
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771360"
---
# <a name="dialog-api-requirement-sets"></a>Ensembles de conditions requises de l’API de dialogue

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’ils nécessitent. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises de l’API de boîte de dialogue, les applications clientes Office qui prennent en charge cet ensemble de conditions requises, ainsi que les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  | Office 2013 sur Windows\*<br>(achat définitif) | Office 2016 ou version ultérieure sur Windows\*<br>(achat définitif)   | Office pour Windows<br>abonnés |  Office sur iPad<br>abonnés  |  Office sur Mac<br>abonnés  | Office sur le web  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1,2  | N/A | N/A | Voir prise en charge<br>section ci-dessous | 2,67 ou version ultérieure | 16,37 ou version ultérieure | Juin 2020 | S/O |
| DialogApi 1.1  | Build 15.0.4855.1000 ou version ultérieure | Build 16.0.4390.1000 ou version ultérieure | Version 1602 (Build 6741.0000) ou version ultérieure | 1.22 ou version ultérieure | 15.20 ou version ultérieure | Janvier 2017 | Version 1608 (Build 7601.6800) ou version ultérieure|

>\* Les utilisateurs du bureau unique peuvent ne pas avoir accepté tous les correctifs et mises à jour. Si c’est le cas, la DLL qu’Office utilise pour signaler sa version dans l’interface utilisateur peut être supérieure aux versions indiquées ici même si les dll mises à jour nécessaires pour prendre en charge DialogApi n’ont pas été installées sur l’ordinateur de l’utilisateur. Pour vous assurer que le correctif nécessaire est installé, l’utilisateur doit accéder à la liste Office Update List ([office 2013 List](/officeupdates/msp-files-office-2013) ou [Office 2016 List](/officeupdates/msp-files-office-2016)), rechercher **osfclient-x-none** et installer le correctif répertorié.

## <a name="office-on-windows-subscription-support"></a>Prise en charge d’Office sur Windows (abonnement)

L’ensemble de conditions requises DialogApi 1,2 est pris en charge dans la version 2005 du canal grand public (Build, 12827,20268 ou version ultérieure). Pour Office sous Windows, la fonctionnalité est également prise en charge dans le canal Semi-Annual et les versions mensuelles de canaux d’entreprise disponibles pour le 9 juin, 2020 ou une version ultérieure. Les versions minimales prises en charge pour chaque canal sont les suivantes :  

|Canal | Version | Build|
|:-----|:-----|:-----|
|Canal actuel | 2005 ou version ultérieure | 12827,20160 ou version ultérieure|
|Canal Entreprise mensuel | 2004 ou version ultérieure | 12730,20430 ou version ultérieure|
|Canal Entreprise semi-annuel | 2002 ou version ultérieure | 12527,20720 ou version ultérieure|

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="dialog-api-11-and-12"></a>API de dialogue 1,1 et 1,2

L’API de boîte de dialogue 1.1 est la première version de l’API. L’ensemble de conditions requises 1,2 ajoute la prise en charge de l’envoi de données à partir de la page parent vers la boîte de dialogue avec la méthode [Office. Dialog. messageChild](/javascript/api/office/office.dialog#messageChild_message_) . Pour plus d’informations sur ces API, consultez la rubrique référence de l' [API Dialog](/javascript/api/office/office.ui) .

## <a name="see-also"></a>Voir aussi

- [Utiliser l’API de boîte de dialogue Office dans les compléments Office](../../develop/dialog-api-in-office-add-ins.md)
- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécifier les applications Office et les exigences de l’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
