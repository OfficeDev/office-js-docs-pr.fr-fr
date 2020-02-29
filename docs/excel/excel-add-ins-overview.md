---
title: Présentation des compléments Excel
description: ''
ms.date: 07/05/2019
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 6f2e319c5de310df5bd30a1161332d03344f0021
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325100"
---
# <a name="excel-add-ins-overview"></a>Présentation des compléments Excel

Un complément Excel vous permet d’étendre les fonctionnalités de l’application Excel sur plusieurs plateformes, notamment Windows, Mac et iPad, ainsi que dans un navigateur web. Utilisez les compléments Excel dans un classeur pour :

- Interagir avec des objets Excel, lire et écrire des données Excel
- Étendre les fonctionnalités à l’aide du volet Office web ou du volet de contenu
- Ajouter des boutons personnalisés au ruban ou des éléments au menu contextuel
- Ajouter des fonctions personnalisées
- Fournir une interaction améliorée à l’aide de la fenêtre de dialogue

La plateforme de compléments Office fournit la structure et les API JavaScript Office.js qui vous permettent de créer et d’exécuter des compléments Excel. En utilisant la plateforme de compléments Office pour créer votre complément Excel, vous bénéficierez des avantages suivants :

* **Prise en charge sur plusieurs plateformes** : les compléments Excel s’exécutent sur Office sur le web, Windows, Mac et iPad.
* **Déploiement centralisé** : les administrateurs peuvent rapidement et facilement déployer des compléments Excel vers les utilisateurs d’une organisation.
* **Utilisation de technologies web standard** : créez votre complément Excel en utilisant des technologies web connues telles qu’HTML, CSS et JavaScript.
* **Distribution via AppSource** : partagez votre complément Excel avec un large public en le publiant sur [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=53245fad-fcbe-41f8-9f97-b0840264f97c&omexanonuid=4a0102fb-b31a-4b9f-9bb0-39d4cc6b789d).

> [!NOTE]
> Les compléments Excel sont différents des compléments COM ou VST, qui sont des solutions d’intégration Office antérieures s’exécutant uniquement sur Office pour Windows. Contrairement aux compléments COM, les compléments Excel ne nécessitent pas l’installation de code sur l’appareil d’un utilisateur ou dans Excel.

## <a name="components-of-an-excel-add-in"></a>Composants d’un complément Excel

Un complément Excel comprend deux composants de base : une application web et un fichier de configuration, appelé fichier manifeste. 

L’application web utilise l’[API JavaScript pour Office](/office/dev/add-ins/reference/javascript-api-for-office) pour interagir avec des objets dans Excel et peut également faciliter l’interaction avec les ressources en ligne. Par exemple, un complément peut effectuer une des opérations suivantes :

* Créer, lire, mettre à jour et supprimer des données dans le classeur (feuilles de calcul, plages, tableaux, graphiques, éléments nommés, etc.).
* Effectuer une autorisation utilisateur avec un service en ligne à l’aide du flux OAuth 2.0 standard.
* Émettre des demandes d’API à Microsoft Graph ou toute autre API.

L’application web peut être hébergée sur un serveur web et peut être créée à l’aide de structures de côté client (par exemple, Angular, React, jQuery) ou des technologies côté serveur (par exemple, ASP.NET, Node.js, PHP).

Le [manifeste](../develop/add-in-manifests.md) est un fichier de configuration XML qui définit la façon dont le complément est intégré dans les clients Office en spécifiant des paramètres et fonctionnalités telles que :

* L’URL de l’application web du complément.
* Le nom d’affichage, la description, l’ID, la version et les paramètres régionaux par défaut du complément.
* La manière dont le complément est intégré à Excel, y compris toute interface utilisateur personnalisée créée par le complément (boutons du ruban, menus contextuels, etc.).
* Les autorisations requises par le complément, comme la lecture du document ou l’écriture dans celui-ci.

Pour permettre aux utilisateurs finals d’installer et d’utiliser un complément Excel, publiez son manifeste dans AppSource ou dans un catalogue de compléments. Pour plus de détails sur la publication dans AppSource, reportez-vous à la rubrique [Mise à disposition de vos solutions dans AppSource et dans Office](/office/dev/store/submit-to-appsource-via-partner-center).

## <a name="capabilities-of-an-excel-add-in"></a>Fonctionnalités d’un complément Excel

En plus d’interagir avec le contenu du classeur, les compléments Excel peuvent ajouter des boutons personnalisés au ruban ou des commandes de menu, insérer des volets Office, ajouter des fonctions personnalisées, ouvrir des boîtes de dialogue et même incorporer des objets web enrichis, tels que des graphiques ou des visualisations interactives dans une feuille de calcul.

### <a name="add-in-commands"></a>Commandes de complément

Les commandes de complément sont des éléments d’interface utilisateur qui étendent l’interface utilisateur Excel et lancent des actions dans votre complément. Vous pouvez utiliser les commandes de complément pour ajouter un bouton au ruban ou un élément à un menu contextuel dans Excel. Lorsque les utilisateurs sélectionnent une commande de complément, ils lancent des actions telles que l’exécution de code JavaScript ou l’affichage d’une page du complément dans un volet Office. 

**Commandes de complément**

![Commandes de complément dans Excel](../images/excel-add-in-commands-script-lab.png)

Pour plus d’informations sur les fonctionnalités des commandes, les plateformes prises en charge et les bonnes pratiques pour le développement de commandes, reportez-vous à la rubrique [Commandes de complément pour Excel, Word et PowerPoint](../design/add-in-commands.md).

### <a name="task-panes"></a>Volets Office

Les volets Office sont des surfaces d’interface qui s’affichent généralement sur le côté droit de la fenêtre dans Excel. Les volets Office permettent aux utilisateurs d’accéder à des contrôles d’interface qui exécutent du code pour modifier le document Excel ou afficher les données d’une source de données. 

**Volet Office**

![Complément du volet Office dans Excel](../images/excel-add-in-task-pane-insights.png)

Pour plus d’informations sur les volets Office, reportez-vous à [Volets Office dans les compléments Office](../design/task-pane-add-ins.md). Pour consulter un exemple qui implémente un volet Office dans Excel, reportez-vous à [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends).

### <a name="custom-functions"></a>Fonctions personnalisées

Les fonctions personnalisées permettent aux développeurs d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément. Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`. 

**Fonction personnalisée**

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

Pour plus d’informations sur les fonctions personnalisées, voir [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md).

### <a name="dialog-boxes"></a>Boîtes de dialogue

Les boîtes de dialogue sont des surfaces qui flottent au-dessus de la fenêtre active de l’application Excel. Vous pouvez utiliser les boîtes de dialogue pour des tâches comme l’affichage de pages de connexion impossibles à ouvrir directement dans un volet Office, les demandes de confirmation d’une action par l’utilisateur ou l’hébergement de vidéos pouvant être trop petites si elles sont limitées à un volet Office. Utilisez l’[API de dialogue](/javascript/api/office/office.ui) pour ouvrir des boîtes de dialogue dans votre complément Excel.

**Boîte de dialogue**

![Boîte de dialogue de complément dans Excel](../images/excel-add-in-dialog-choose-number.png)

Pour plus d’informations sur les boîtes de dialogue et l’API de dialogue, reportez-vous aux rubriques [Boîtes de dialogue dans les compléments Office](../design/dialog-boxes.md) et [Utiliser l’API de dialogue dans vos compléments Office](../develop/dialog-api-in-office-add-ins.md).

### <a name="content-add-ins"></a>Compléments de contenu

Les compléments de contenu sont des surfaces que vous pouvez incorporer directement dans les documents Excel. Vous pouvez utiliser des compléments de contenu pour incorporer des objets riches, basés sur le web, tels que des graphiques, des visualisations de données ou des supports dans une feuille de calcul, ou autoriser l’accès des utilisateurs aux options d’interface qui exécutent le code pour modifier le document Excel ou afficher des données à partir d’une source de données. Utilisez les compléments de contenu lorsque vous souhaitez incorporer des fonctionnalités directement dans le document.

**Complément de contenu**

![Complément de contenu dans Excel](../images/excel-add-in-content-map.png)

Pour plus d’informations sur les compléments de contenu, reportez-vous à [Compléments Office de contenu](../design/content-add-ins.md). Pour consulter un exemple qui implémente un complément de contenu dans Excel, reportez-vous à [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) dans GitHub.

## <a name="javascript-apis-to-interact-with-workbook-content"></a>API JavaScript permettant d’interagir avec le contenu du classeur

Un complément Excel interagit avec des objets dans Excel en utilisant l’[API Office JavaScript](/office/dev/add-ins/reference/javascript-api-for-office), qui inclut deux modèles d’objets JavaScript :

* **API JavaScript pour Excel** : incluse dans Office 2016, l’[API JavaScript pour Excel](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) fournit des objets Excel fortement typés que vous pouvez utiliser pour accéder aux feuilles de calcul, aux plages, aux tableaux, aux graphiques et bien plus encore. 

* **API commune** : incluse dans Office 2013, l’API commune vous permet d’accéder à des fonctionnalités telles qu’une interface utilisateur, des boîtes de dialogue et des paramètres du client, qui sont communes à plusieurs types d’applications Office. Étant donné que l’API commune fournit des fonctionnalités limitées pour une interaction avec Excel, vous pouvez l’utiliser si votre complément doit s’exécuter sur Excel 2013.

## <a name="next-steps"></a>Étapes suivantes

Apprenez à [créer votre premier complément Excel](../quickstarts/excel-quickstart-jquery.md). Découvrez ensuite les [concepts fondamentaux](excel-add-ins-core-concepts.md) de la création de compléments Excel.

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
- [Création de compléments Office](../overview/office-add-ins-fundamentals.md)
- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Référence sur l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)