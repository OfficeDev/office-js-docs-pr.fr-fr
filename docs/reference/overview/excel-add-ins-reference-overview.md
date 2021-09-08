---
title: Vue d’ensemble de l’API JavaScript pour Excel
description: En savoir plus sur l’API JavaScript pour Excel
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 80340b4990b56b2ba4d51f2a028480af3e267828
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937258"
---
# <a name="excel-javascript-api-overview"></a>Vue d’ensemble de l’API JavaScript pour Excel

Un complément Excel interagit avec des objets dans Excel en utilisant l’API JavaScript pour Office, qui inclut deux modèles d’objets JavaScript :

* **API JavaScript Excel** : il s’agit des [API spécifiques à l’application](../../develop/application-specific-api-model.md) pour Excel. Introduite avec Office 2016, L’[API JavaScript Excel](/javascript/api/excel) fournit des objets fortement typés que vous pouvez utiliser pour accéder aux feuilles de calcul, plages, tableaux, graphiques, etc.

* **API communes** : incluses dans Office 2013, les [API communes](/javascript/api/office) peuvent être utilisées pour accéder à des fonctionnalités telles qu’une interface utilisateur, des boîtes de dialogue et des paramètres du client, qui sont communes à plusieurs types d’applications Office.

Cette section de la documentation se concentre sur l’API JavaScript Excel, que vous allez utiliser pour développer la majorité des fonctionnalités des compléments qui ciblent Excel sur le web ou Excel 2016 ou version ultérieure. Pour plus d’informations sur l’API commune, consultez [Modèle objet d’API JavaScript commun](../../develop/office-javascript-api-object-model.md).

## <a name="learn-object-model-concepts"></a>Découvrir les concepts du modèle d’objet

Voir [Modèle d’objet JavaScript Excel dans les compléments Office](../../excel/excel-add-ins-core-concepts.md) pour plus d’informations sur les concepts importants du modèle d’objet.

Pour apprendre à utiliser l’API JavaScript pour Excel afin d’accéder à des objets dans Excel, suivez le [didacticiel sur les compléments Excel](../../tutorials/excel-tutorial.md).

## <a name="learn-api-capabilities"></a>En savoir plus sur les fonctionnalités des API

Chaque fonctionnalité principale de l’API Excel inclut un article ou un ensemble d’articles sur la façon dont cette fonctionnalité et le modèle d’objet approprié sont utilisés.

* [Graphiques](../../excel/excel-add-ins-charts.md)
* [Commentaires](../../excel/excel-add-ins-comments.md)
* [Mise en forme conditionnelle](../../excel/excel-add-ins-conditional-formatting.md)
* [Fonctions personnalisées](../../excel/custom-functions-overview.md)
* [Validation des données](../../excel/excel-add-ins-data-validation.md)
* [Événements](../../excel/excel-add-ins-events.md)
* [PivotTables](../../excel/excel-add-ins-pivottables.md)
* [Plages](../../excel/excel-add-ins-ranges-get.md) et [Cellules](../../excel/excel-add-ins-cells.md)
* [RangeAreas (Plages multiples)](../../excel/excel-add-ins-multiple-ranges.md)
* [Formes](../../excel/excel-add-ins-shapes.md)
* [Tableaux](../../excel/excel-add-ins-tables.md)
* [Classeurs et API au niveau de l’application](../../excel/excel-add-ins-workbooks.md)
* [Feuilles de calcul](../../excel/excel-add-ins-worksheets.md)

Pour en savoir plus sur le modèle objet de l’API JavaScript pour Excel, consultez la [documentation de référence sur l’API JavaScript pour Excel](/javascript/api/excel).

## <a name="try-out-code-samples-in-script-lab"></a>Tester les exemples de code dans Script Lab

Utilisez [Script Lab](../../overview/explore-with-script-lab.md) pour commencer rapidement avec une collection d’exemples intégrés qui vous explique comment accomplir des tâches avec l’API. Vous pouvez exécuter les exemples dans Script Lab de manière à afficher instantanément le résultat dans le volet Office ou la feuille de calcul, examiner les exemples pour découvrir le fonctionnement de l’API, voire utiliser les exemples pour prototyper votre propre complément.

## <a name="see-also"></a>Voir aussi

* [Documentation sur les compléments Excel](../../excel/index.yml)
* [Présentation des compléments Excel](../../excel/excel-add-ins-overview.md)
* [Référence de l’API JavaScript pour Excel](/javascript/api/excel)
* [Application cliente Office et disponibilité de la plateforme pour les compléments Office](../../overview/office-add-in-availability.md)
* [Utilisation du modèle API propre à l’application](../../develop/application-specific-api-model.md)
