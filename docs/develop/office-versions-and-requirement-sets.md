---
title: Versions d’Office et ensembles de conditions requises
description: Plateformes Office.js prises en charge à l'aide de l'API JavaScript.
ms.date: 07/08/2021
ms.localizationpriority: high
ms.openlocfilehash: 62dbac6c5f08596792c29ce5451cdaca077ddcdb
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/26/2022
ms.locfileid: "64483658"
---
# <a name="office-versions-and-requirement-sets"></a>Versions d’Office et ensembles de conditions requises

Il existe de nombreuses versions d’Office sur plusieurs plateformes ; celles-ci ne prennent pas forcément en charge toutes les API dans l’interface API JavaScript pour Office (Office.js). Vous n’avez pas toujours le contrôle sur la version d’Office que vos utilisateurs ont installée.  Pour gérer cette situation, nous fournissons un système nommé ensembles de conditions requises pour vous aider à déterminer si une application Office prend en charge les fonctionnalités dont vous avez besoin dans votre complément Office.

> [!NOTE]
>
> - Office s’exécute sur plusieurs plateformes, y compris sur Windows, dans un navigateur, un Mac et un iPad.
> - Parmi les applications Office, voici quelques exemples de produits Office : Excel, Word, PowerPoint, Outlook, OneNote et autres.  
> - Un ensemble de conditions requises est un groupe nommé de membres d’API, par exemple : `ExcelApi 1.5`, `WordApi 1.3`, et ainsi de suite.  

## <a name="how-to-check-your-office-version"></a>Vérification de votre version d’Office

Pour identifier la version d’Office que vous utilisez, à partir d’une application Office, sélectionnez le menu **Fichier**, puis sélectionnez **Compte**. La version d’Office s’affiche dans la section **Informations sur le produit**. Par exemple, la capture d’écran suivante indique la version 1802 d’Office (build 9026.1000).

![Vérifier la version de votre Office.](../images/office-version.png)

## <a name="office-requirement-sets-availability"></a>Disponibilité des ensembles de conditions requises Office

Les compléments Office peuvent utiliser des ensembles de conditions requises d’API pour déterminer si l’application Office prend en charge les membres d’API nécessaires. La prise en charge des ensembles de conditions requises varie selon l’application Office et la version de cette dernière (voir la section précédente).

Certaines applications Office ont leurs propres ensembles de conditions requises d’API. Par exemple, le premier ensemble de conditions requises pour l’API Excel était `ExcelApi 1.1` et le premier ensemble de conditions requises pour l’API Word était `WordApi 1.1`. Depuis lors, de nombreux ensembles de conditions requises ExcelApi et WordApi ont été ajoutés pour proposer des fonctionnalités d’API supplémentaires.

En outre, d’autres fonctionnalités telles que les commandes de complément (extensibilité du ruban) et la possibilité de lancer des boîtes de dialogue (API de boîte de dialogue) ont été ajoutées à l’API commune. Les commandes de complément et les ensembles de conditions requises d’API de boîte de dialogue sont des exemples d’ensembles d’API que les différentes applications Office partagent en commun.

Un complément peut utiliser uniquement des API dans les ensembles de conditions requises qui sont prises en charge par la version de l’application Office sur laquelle le complément est exécuté. Pour savoir exactement quels ensembles de conditions requises sont disponibles pour une version spécifique de l’application Office, reportez-vous aux articles suivants sur les ensembles de conditions requises propres aux applications.

- [Ensembles de conditions requises de l’API JavaScript pour Excel](/javascript/api/requirement-sets/excel-api-requirement-sets) (ExcelApi)
- [Ensembles de conditions requises de l’API JavaScript pour Word](/javascript/api/requirement-sets/word-api-requirement-sets) (WordApi)
- [Ensembles de conditions requises de l’API JavaScript pour OneNote](/javascript/api/requirement-sets/onenote-api-requirement-sets) (OneNoteApi)
- [Ensembles de conditions requises de l’API JavaScript pour PowerPoint](/javascript/api/requirement-sets/powerpoint-api-requirement-sets) (PowerPointApi)
- [Présentation de l’ensemble de conditions requises pour les API Outlook](/javascript/api/requirement-sets/outlook-api-requirement-sets) (Mailbox)

Certains ensembles de conditions requises contiennent des API qui peuvent être utilisées par n’importe quelle application Office. Pour plus d’informations sur ces ensembles de conditions requises, reportez-vous aux articles suivants.

- [Ensembles de conditions requises communes pour Office](/javascript/api/requirement-sets/office-add-in-requirement-sets)
- [Ensembles de conditions requises concernant les commandes de complément](/javascript/api/requirement-sets/add-in-commands-requirement-sets)
- [Ensembles de conditions requises de l’API de boîte de dialogue](/javascript/api/requirement-sets/dialog-api-requirement-sets)
- [Ensembles de conditions requises de l’API d’identité](/javascript/api/requirement-sets/identity-api-requirement-sets)

Le numéro de version d’un ensemble de conditions requises, par exemple « 1.1 » dans `ExcelApi 1.1`, est défini par rapport à l’application Office. Le numéro de version d’un ensemble donné de conditions requises (par exemple, `ExcelApi 1.1`) ne correspond pas au numéro de version d’Office.js, ni aux ensembles de conditions requises pour d’autres applications Office (comme Word, Outlook, etc.).  Les ensembles de conditions requises pour les différentes applications Office sont publiés à des rythmes différents. Par exemple, `ExcelApi 1.5` a été publié avant l’ensemble de conditions requises `WordApi 1.3`.

La bibliothèque d’API JavaScript Office (Office.js) inclut tous les ensembles de conditions requises actuellement disponibles. Bien qu’il existe un ensemble de conditions requises `ExcelApi 1.3` et `WordApi 1.3`, il n’existe aucun ensemble de conditions requises `Office.js 1.3` . La dernière version d’Office.js est conservée en tant que point de terminaison Office unique fourni via le réseau de distribution de contenu (CDN). Pour plus d’informations sur le CDN Office.js, notamment la façon dont le contrôle de version et la compatibilité descendante sont gérés, consultez [Présentation de l’API JavaScript Office](../develop/understanding-the-javascript-api-for-office.md).

## <a name="specify-office-applications-and-requirement-sets"></a>Spécifier les ensembles de conditions requises et les applications Office

Il existe différentes méthodes pour spécifier les applications Office et les ensembles de conditions qui sont requis par un complément.  Pour plus d’informations, consultez la rubrique [Spécifier les applications Office et les conditions requises d’API](../develop/specify-office-hosts-and-api-requirements.md).

## <a name="see-also"></a>Voir aussi

- [Spécifier les exigences en matière d’applications Office et d’API](../develop/specify-office-hosts-and-api-requirements.md)
- [Installer la dernière version d’Office](../develop/install-latest-office-version.md)
- [Aperçu des canaux de mise à jour pour les applications Microsoft 365](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Réinventez la productivité avec Microsoft 365 et Microsoft Teams](https://products.office.com/compare-all-microsoft-office-products?tab=2)
