---
title: Versions d’Office et ensembles de conditions requises
description: Plateformes Office.js prises en charge à l'aide de l'API JavaScript.
ms.date: 09/14/2022
ms.localizationpriority: high
ms.openlocfilehash: 669977f87974a1ec5519ddbbe3d38c5a290ec84f
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/30/2022
ms.locfileid: "68234906"
---
# <a name="office-versions-and-requirement-sets"></a>Versions d’Office et ensembles de conditions requises

Il existe de nombreuses versions d’Office sur plusieurs plateformes, celles-ci ne prenant pas forcément en charge toutes les API dans l’interface API JavaScript pour Office (Office.js). Office 2013 sur Windows était la première version d’Office qui prenait en charge les compléments Office. Vous n’avez peut-être pas toujours le contrôle sur la version d’Office que vos utilisateurs ont installée. Pour gérer cette situation, nous fournissons un système appelé ensembles de conditions requises pour vous aider à déterminer si une application Office prend en charge les fonctionnalités dont vous avez besoin dans votre complément Office.

> [!NOTE]
>
> - Office s’exécute sur plusieurs plateformes, y compris sur Windows, dans un navigateur, un Mac et un iPad.
> - Les produits Office sont des exemples d’applications Office : Excel, Word, PowerPoint, Outlook, OneNote, etc.
> - Office est disponible par abonnement Microsoft 365 ou licence perpétuelle. La version perpétuelle est disponible par contrat de licence en volume ou par vente au détail.
> - Un ensemble de conditions requises est un groupe nommé de membres d’API, par exemple, `ExcelApi 1.5`, `WordApi 1.3`et ainsi de suite.

## <a name="how-to-check-your-office-version"></a>Vérification de votre version d’Office

Pour identifier la version d’Office que vous utilisez, à partir d’une application Office, sélectionnez le menu **Fichier**, puis sélectionnez **Compte**. La version d’Office apparaît dans la section **Informations sur le produit** . Par exemple, la capture d’écran suivante indique Office version 1802 (build 9026.1000).

![Vérifier la version de votre Office.](../images/office-version.png)

> [!NOTE]
> Si votre version d’Office est différente de celle-ci, consultez [Quelle version d’Outlook ai-je ?](https://support.microsoft.com/office/b3a9568c-edb5-42b9-9825-d48d82b2257c) ou [À propos d’Office : quelle version d’Office utilise-t-on ?](https://support.microsoft.com/topic/932788b8-a3ce-44bf-bb09-e334518b8b19) pour comprendre comment obtenir ces informations pour votre version.

## <a name="office-requirement-sets-availability"></a>Disponibilité des ensembles de conditions requises Office

Les compléments Office peuvent utiliser des ensembles de conditions requises d’API pour déterminer si l’application Office prend en charge les membres de l’API qu’elle doit utiliser. La prise en charge de l’ensemble de conditions requises varie selon l’application Office et la version de l’application Office (voir la section précédente [Comment vérifier votre version d’Office](#how-to-check-your-office-version)).

Some Office applications have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.

Par ailleurs, d’autres fonctionnalités telles que les commandes de complément (extensibilité du ruban) et la possibilité de lancer des boîtes de dialogue (API de boîte de dialogue) ont été ajoutées à l’API commune. Les commandes de complément et les ensembles de conditions requises d’API de boîte de dialogue sont des exemples d’ensembles d’API que différentes applications Office partagent en commun.

An add-in can only use APIs in requirement sets that are supported by the version of Office application where the add-in is running. To know exactly which requirement sets are available for a specific Office application version, refer to the following application-specific requirement set articles.

- [Ensembles de conditions requises de l’API JavaScript pour Excel](/javascript/api/requirement-sets/excel/excel-api-requirement-sets) (ExcelApi)
- [Ensembles de conditions requises de l’API JavaScript pour OneNote](/javascript/api/requirement-sets/onenote/onenote-api-requirement-sets) (OneNoteApi)
- [Ensembles de conditions requises de l’API JavaScript Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets) (boîte aux lettres)
- [Ensembles de conditions requises de l’API JavaScript pour PowerPoint](/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets) (PowerPointApi)
- [Ensembles de conditions requises de l’API JavaScript pour Word](/javascript/api/requirement-sets/word/word-api-requirement-sets) (WordApi)

Certains ensembles de conditions requises contiennent des API qui peuvent être utilisées par plusieurs applications Office. Pour plus d’informations sur ces ensembles de conditions requises, reportez-vous aux articles suivants.

- [Ensembles de conditions requises communes pour Office](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
- [Ensembles de conditions requises concernant les commandes de complément](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets)
- [Ensembles de conditions requises de l’API de boîte de dialogue](/javascript/api/requirement-sets/common/dialog-api-requirement-sets)
- [Ensembles de conditions requises d’origine de boîte de dialogue](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets)
- [Ensembles de conditions requises de l’API d’identité](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
- [Ensembles de conditions requises de coercition d’image](/javascript/api/requirement-sets/common/image-coercion-requirement-sets)
- [Ensembles de conditions requises pour les raccourcis clavier](/javascript/api/requirement-sets/common/keyboard-shortcuts-requirement-sets)
- [Séries de conditions requises pour ouvrir une fenêtre de navigateur](/javascript/api/requirement-sets/common/open-browser-window-api-requirement-sets)
- [Ensembles de conditions requises des API ruban](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets)
- [Ensembles de conditions requises d'exécution partagés](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)

The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office application. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office applications (e.g., Word, Outlook, etc.).  Requirement sets for the different Office applications are released at different rates. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.

The Office JavaScript API library (Office.js) includes all requirement sets that are currently available. While there is such a thing as requirement sets `ExcelApi 1.3` and `WordApi 1.3`, there is no `Office.js 1.3` requirement set. The latest release of Office.js is maintained as a single Office endpoint delivered via the content delivery network (CDN). For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Understanding the Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md).

## <a name="specify-office-applications-and-requirement-sets"></a>Spécifier les ensembles de conditions requises et les applications Office

There are various ways to specify which Office applications and requirement sets are required by an add-in.  For detailed information, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md)

## <a name="see-also"></a>Voir aussi

- [Spécifier les exigences en matière d’applications Office et d’API](../develop/specify-office-hosts-and-api-requirements.md)
- [Installer la dernière version d’Office](../develop/install-latest-office-version.md)
- [Aperçu des canaux de mise à jour pour les applications Microsoft 365](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Réinventez la productivité avec Microsoft 365 et Microsoft Teams](https://products.office.com/compare-all-microsoft-office-products?tab=2)
