---
title: Versions d’Office et ensembles de conditions requises
description: ''
ms.date: 03/29/2018
ms.openlocfilehash: ac3ae4fa3eeca9cfbd56b15168fc39d67139680d
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505992"
---
# <a name="office-versions-and-requirement-sets"></a>Versions d’Office et ensembles de conditions requises

Il existe plusieurs versions d’Office sur plusieurs plateformes, et elles ne tous prennent en charge toutes les API dans l’API JavaScript pour Office (Office.js). Vous n’avez pas toujours contrôler la version d’Office, vos utilisateurs ont installé.  Pour gérer cette situation, nous fournissons un système appelé ensembles de ressources pour vous aider à déterminer si un hôte Office prend en charge les fonctionnalités que vous avez besoin dans votre complément Office. 

> [!NOTE]
> - Office peut être exécuté sur plusieurs plateformes, notamment Office pour Windows, Office Online, Office pour Mac et Office pour iPad.  
> - Parmi les hôtes Office, voici quelques exemples de produits Office : Excel, Word, PowerPoint, Outlook, OneNote et autres.  
> - Un ensemble de conditions requises est un groupe nommé de membres d’API, par exemple : `ExcelApi 1.5`, `WordApi 1.3`, et ainsi de suite.  


## <a name="how-to-check-your-office-version"></a>Vérification de votre version d’Office

Pour identifier la version d’Office que vous utilisez, à partir d’au sein d’une application Office, sélectionnez le menu **fichier** , puis cliquez sur **compte**. La version d’Office s’affiche dans la section **Informations sur le produit** . Par exemple, la capture d’écran suivante indique Office Version 1802 (Build 9026.1000) :

![Vérification de votre version d’Office](../images/office-version-number-ui.jpg)


## <a name="office-requirement-sets-availability"></a>Disponibilité des ensembles de conditions requises Office

Compléments Office permet de déterminer si l’hôte Office prend en charge les membres de l’API il faut utiliser ensembles d’API. Prise en charge du jeu requise varie en fonction hôte Office et la version d’hôte Office (voir la section précédente).

Certains hôtes Office ont leurs propres ensembles d’API. Par exemple, la première condition définie pour l’API Excel a `ExcelApi 1.1` et de la première demande défini pour l’API de Word a été `WordApi 1.1`. Depuis, plusieurs nouveaux ensembles ExcelApi et ensembles WordApi ont été ajoutées pour fournir des fonctionnalités supplémentaires API.

En outre, les autres fonctionnalités telles que les commandes de complément (extensibilité du ruban) et la possibilité de lancer des boîtes de dialogue (boîte de dialogue API) ont été ajoutés à l’API courantes. Dans Ajouter des commandes et des ensembles d’API de boîte de dialogue sont des exemples de jeux d’API que les hôtes Office différents ont en commun.

Un complément peut uniquement utiliser API dans les ensembles sont prises en charge par la version de l’hôte Office où le complément est en cours d’exécution. Pour savoir exactement quels ensembles sont disponibles pour une version d’hôte Office spécifique, reportez-vous aux articles set exigence spécifique à l’hôte suivantes :

- [Ensembles de conditions requises de l’API JavaScript pour Excel](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js) (ExcelApi)
- [Ensembles de conditions requises de l’API JavaScript pour Word](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets?view=office-js) (WordApi)
- [Ensembles de conditions requises de l’API JavaScript pour OneNote](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets?view=office-js) (OneNoteApi)
- [Présentation de l’ensemble de conditions requises pour les API Outlook](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets?view=office-js) (MailBox)

Certains ensembles contiennent les API qui peut être utilisé par n’importe quel hôte Office. Pour plus d’informations sur ces ensembles, consultez les articles suivants :

- [Ensembles de conditions requises communes pour Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js)
- [Ensembles de conditions requises concernant les commandes de complément](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets?view=office-js)
- [Ensembles de conditions requises de l’API de boîte de dialogue](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [Ensembles de conditions requises de l’API d’identité](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)

Définie le numéro de version d’une spécification, telles que le « 1.1 » dans `ExcelApi 1.1`, est relative à l’hôte Office. Le numéro de version d’un ensemble donné d’exigence (par exemple, `ExcelApi 1.1`) ne correspond pas au numéro de version d’Office.js ou aux ensembles de conditions pour les autres hôtes Office (par exemple, Word, Outlook, etc.).  Ensembles de ressources pour les hôtes Office différents sont publiés à des dates et des différentes vitesses. Par exemple, `ExcelApi 1.5` a été publié avant la `WordApi 1.3` ensemble de conditions requises.

L’API JavaScript pour la bibliothèque Office (Office.js) inclut tous les ensembles qui sont actuellement disponibles. S’il existe une telle chose comme ensembles `ExcelApi 1.3` et `WordApi 1.3`, il est sans `Office.js 1.3` ensemble de conditions requises. La dernière version d’Office.js est conservée en tant qu’un seul point de terminaison Office remis via le réseau de distribution de contenu (CDN). Pour plus d’informations autour du CDN Office.js, notamment comment le contrôle de version et la compatibilité descendante est géré, voir [Présentation de l’API JavaScript pour Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

## <a name="specify-office-hosts-and-requirement-sets"></a>Spécification des ensembles de conditions requises et des hôtes Office

Il existe différentes façons de spécifier les hôtes Office et les ensembles requis par un complément.  Pour plus d’informations, voir [hôtes Office spécifier et exigences de l’API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)


## <a name="see-also"></a>Voir aussi

- [Spécification des exigences en matière d’hôtes Office et d’API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Installer la dernière version d’Office](https://docs.microsoft.com/office/dev/add-ins/develop/install-latest-office-version)
- [Présentation des canaux de mise à jour pour Office 365 ProPlus](https://docs.microsoft.com/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Tirez le meilleur parti d’Office avec Office 365](https://products.office.com/compare-all-microsoft-office-products?tab=2)
