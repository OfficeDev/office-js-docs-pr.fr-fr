---
title: Ensembles de conditions requises de l’API de dialogue
description: En savoir plus sur les ensembles de conditions requises de l’API de dialogue.
ms.date: 10/05/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 4802189b0dbde30d0d9058b542c35cac47074998
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138554"
---
# <a name="dialog-api-requirement-sets"></a>Ensembles de conditions requises de l’API de dialogue

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’ils nécessitent. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises de l’API de dialogue, les applications clientes Office qui la prise en charge, ainsi que les numéros de build ou de version de l’application Office.

| Ensemble de conditions requises | Office 2013 sur Windows\*<br>(achat définitif) | Office 2016 sur Windows\*<br>(achat définitif) | Office 2019 sur Windows\*<br>(achat définitif) | Office 2021 ou une Windows\*<br>(achat définitif) | Office pour Windows<br>(abonnement) | Office sur iPad<br>(abonnement) |  Office sur Mac<br>(abonnement) | Office sur le web | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.2  | N/A | N/A | N/A | Build 16.0.14326.20454 ou ultérieure | Voir la prise en charge<br>section ci-dessous | 2.37 ou ultérieure | 16.37 ou ultérieure | Juin 2020 | S/O |
| DialogApi 1.1  | Build 15.0.4855.1000 ou version ultérieure | Build 16.0.4390.1000 ou version ultérieure | Build 16.0.12527.20720 ou ultérieure | Build 16.0.14326.20454 ou ultérieure | Version 1602 (Build 6741.0000) ou version ultérieure | 1.22 ou version ultérieure | 15.20 ou version ultérieure | Janvier 2017 | Version 1608 (Build 7601.6800) ou version ultérieure|

>\*Les utilisateurs de l’achat Office n’ont peut-être pas accepté tous les correctifs et mises à jour. Si c’est le cas, la DLL que Office utilise pour signaler sa version dans l’interface utilisateur peut être supérieure aux versions répertoriées ici, même si les DLL mises à jour nécessaires pour prendre en charge DialogApi n’ont pas été installées sur l’ordinateur de l’utilisateur. Pour s’assurer que le correctif nécessaire est installé, l’utilisateur doit se rendre dans la liste des mises à jour Office ([liste Office 2013](/officeupdates/msp-files-office-2013) ou [Office 2016](/officeupdates/msp-files-office-2016)), rechercher **osfclient-x-none** et installer le correctif répertorié.

## <a name="office-on-windows-subscription-support"></a>Office prise en charge Windows (abonnement)

L’ensemble de conditions requises DialogApi 1.2 est pris en charge dans le Canal consommateur version 2005 (build 12827.20268 ou version supérieure). Pour Office sur Windows, la fonctionnalité est également prise en charge dans les builds du canal Semi-Annual et du canal Enterprise mensuel disponibles le 9 juin 2020 ou une date ultérieure. Les builds minimales prise en charge pour chaque canal sont les suivantes :  

|Canal | Version | Build|
|:-----|:-----|:-----|
|Canal actuel | 2005 ou supérieure | 12827.20160 ou supérieur|
|Canal mensuel des entreprises | 2004 ou supérieure | 12730.20430 ou supérieur|
|Canal Entreprise semestriel | 2002 ou supérieure | 12527.20720 ou supérieur|

## <a name="office-versions-and-build-numbers"></a>Numéros de version et de build d’Office

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Présentation d’Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="dialog-api-11-and-12"></a>API de boîte de dialogue 1.1 et 1.2

L’API de boîte de dialogue 1.1 est la première version de l’API. L’ensemble de conditions requises 1.2 ajoute la prise en charge de l’envoi de données à partir de la page parent à la boîte de dialogue à l’aide de [la méthode Office.dialog.messageChild.](/javascript/api/office/office.dialog#messageChild_message_) Pour plus d’informations sur ces API, voir la rubrique de référence [de l’API](/javascript/api/office/office.ui) de dialogue.

## <a name="see-also"></a>Voir aussi

- [Utiliser l’API de boîte de dialogue Office dans les compléments Office](../../develop/dialog-api-in-office-add-ins.md)
- [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md)
- [Spécifier les exigences en matière d’applications Office et d’API](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../develop/add-in-manifests.md)
