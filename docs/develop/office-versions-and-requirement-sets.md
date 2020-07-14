---
title: Versions d’Office et ensembles de conditions requises
description: Plateformes Office.js prises en charge à l'aide de l'API JavaScript
ms.date: 07/07/2020
localization_priority: Priority
ms.openlocfilehash: 02f3d91256ea05e526ebe2e4e4090b1908d7292a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093580"
---
# <a name="office-versions-and-requirement-sets"></a>Versions d’Office et ensembles de conditions requises

There are many versions of Office on several platforms, and they don't all support every API in Office JavaScript API (Office.js). You may not always have control over the version of Office your users have installed.  To handle this situation, we provide a system called requirement sets to help you determine whether an Office host supports the capabilities you need in your Office Add-in. 

> [!NOTE]
> - Office s’exécute sur plusieurs plateformes, y compris sur Windows, dans un navigateur, un Mac et un iPad.
> - Parmi les hôtes Office, voici quelques exemples de produits Office : Excel, Word, PowerPoint, Outlook, OneNote et autres.  
> - Un ensemble de conditions requises est un groupe nommé de membres d’API, par exemple : `ExcelApi 1.5`, `WordApi 1.3`, et ainsi de suite.  

## <a name="how-to-check-your-office-version"></a>Vérification de votre version d’Office

To identify the Office version that you're using, from within an Office application, select the **File** menu, and then choose **Account**. The version of Office will appear in the **Product Information** section. For example, the following screenshot indicates Office Version 1802 (Build 9026.1000):

![Vérification de votre version d’Office](../images/office-version.png)

## <a name="office-requirement-sets-availability"></a>Disponibilité des ensembles de conditions requises Office

Office Add-ins can use API requirement sets to determine whether the Office host supports the API members that it need to use. Requirement set support varies by Office host and the Office host version (see previous section).

Some Office hosts have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.

Par ailleurs, d’autres fonctionnalités telles que les commandes de complément (extensibilité du ruban) et la possibilité de lancer des boîtes de dialogue (API de boîte de dialogue) ont été ajoutées à l’API commune. Les commandes de complément et les ensembles de conditions requises d’API de boîte de dialogue sont des exemples d’ensembles de conditions requises d’API que les différents hôtes Office ont en commun.

An add-in can only use APIs in requirement sets that are supported by the version of Office host where the add-in is running. To know exactly which requirement sets are available for a specific Office host version, refer to the following host-specific requirement set articles:

- [Ensembles de conditions requises de l’API JavaScript pour Excel](../reference/requirement-sets/excel-api-requirement-sets.md) (ExcelApi)
- [Ensembles de conditions requises de l’API JavaScript pour Word](../reference/requirement-sets/word-api-requirement-sets.md) (WordApi)
- [Ensembles de conditions requises de l’API JavaScript pour OneNote](../reference/requirement-sets/onenote-api-requirement-sets.md) (OneNoteApi)
- [Ensembles de conditions requises de l’API JavaScript pour PowerPoint](../reference/requirement-sets/powerpoint-api-requirement-sets.md) (PowerPointApi)
- [Présentation de l’ensemble de conditions requises pour les API Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md) (Mailbox)

Some requirement sets contain APIs that can be used by any Office host. For information about these requirement sets, refer to the following articles:

- [Ensembles de conditions requises communes pour Office](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Ensembles de conditions requises concernant les commandes de complément](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [Ensembles de conditions requises de l’API de boîte de dialogue](../reference/requirement-sets/dialog-api-requirement-sets.md)
- [Ensembles de conditions requises de l’API d’identité](../reference/requirement-sets/identity-api-requirement-sets.md)

The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office host. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office hosts (e.g., Word, Outlook, etc.).  Requirement sets for the different Office hosts are released at different speeds and times. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.

L’API JavaScript pour la bibliothèque Office (Office.js) inclut tous les ensembles de conditions requises actuellement disponibles. Alors qu’il existe des ensembles de conditions requises `ExcelApi 1.3` et `WordApi 1.3`, il n’existe pas d’ensemble de conditions requises `Office.js 1.3`. La dernière version d’Office.js est gérée comme un point de terminaison Office unique remis via le réseau de distribution de contenu (CDN). Pour plus d’informations sur le CDN Office.js, notamment sur la gestion des versions et de la compatibilité avec les anciennes versions, reportez-vous à l’article [Présentation de l’API Interface JavaScript pour Office](../develop/understanding-the-javascript-api-for-office.md).

## <a name="specify-office-hosts-and-requirement-sets"></a>Spécification des ensembles de conditions requises et des hôtes Office

There are various ways to specify which Office hosts and requirement sets are required by an add-in.  For detailed information, see [Specify Office hosts and API requirements](../develop/specify-office-hosts-and-api-requirements.md)

## <a name="see-also"></a>Voir aussi

- [Spécification des exigences en matière d’hôtes Office et d’API](../develop/specify-office-hosts-and-api-requirements.md)
- [Installer la dernière version d’Office](../develop/install-latest-office-version.md)
- [Aperçu des canaux de mise à jour pour les applications Microsoft 365](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Tirez le meilleur parti d’Office avec Office 365](https://products.office.com/compare-all-microsoft-office-products?tab=2)
