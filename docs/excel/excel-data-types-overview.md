---
title: Vue d’ensemble des types de données dans les compléments Excel
description: Les types de données dans l’API JavaScript Excel permettent aux développeurs de compléments Office de travailler avec des valeurs numériques, des images web, des valeurs d’entité, des tableaux mis en forme au sein des valeurs d’entité et des erreurs améliorées en tant que types de données.
ms.date: 12/08/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: f6545eb314ec5aef30a595b755e3f702a3e5df20
ms.sourcegitcommit: ddb1d85186fd6e77d732159430d20eb7395b9a33
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/10/2021
ms.locfileid: "61406668"
---
# <a name="overview-of-data-types-in-excel-add-ins-preview"></a>Vue d’ensemble des types de données dans les compléments Excel (préversion)

> [!NOTE]
> Les API de types de données ne sont actuellement disponibles que dans la préversion publique. L’aperçu API peut être modifiés et n’est pas destinés à utiliser dans un environnement de production. Nous vous recommandons de les tester uniquement dans les environnements de test et de développement. N’utilisez pas un aperçu d’API dans un environnement de production ou dans les documents commerciaux importants.
>
> Pour utiliser les API disponibles en préversion :
>
> - Vous devez référencer la bibliothèque **bêta** sur le CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). Le [fichier de définition de](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) type pour la compilation et la IntelliSense TypeScript se trouve aux CDN et [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Vous pouvez installer ces types avec `npm install --save-dev @types/office-js-preview` . Pour plus d’informations, voir le @microsoft du package NPM [office-js.](https://www.npmjs.com/package/@microsoft/office-js)
> - Vous devrez peut-être rejoindre [Office programme Insider pour](https://insider.office.com) accéder à des builds Office plus récentes.
>
> Pour tester les types de données dans Office sur Windows, vous devez avoir un numéro de build Excel supérieur ou égal à 16.0.14626.10000. Pour tester les types de données dans Office sur Mac, vous devez avoir un numéro de build Excel supérieur ou égal à 16.55.21102600.

Les types de données dans l’API JavaScript Excel permettent aux développeurs de compléments d’organiser des structures de données complexes en tant qu’objets, tels que des valeurs numériques, des images web et des valeurs d’entité mises en forme.

Avant l’ajout des types de données, l’API JavaScript Excel prenait en charge les types de données chaîne, nombre, booléen et d’erreur. La couche de mise en forme de l’interface utilisateur Excel est capable d’ajouter des devises, des dates et d’autres types de mise en forme aux cellules qui contiennent les quatre types de données d’origine, mais cette couche de mise en forme contrôle uniquement l’affichage des types de données d’origine dans l’interface utilisateur Excel. La valeur du nombre sous-jacent n’est pas modifiée, même lorsqu’une cellule de l’interface utilisateur Excel est mise en forme en tant que devise ou date. Cet écart entre une valeur sous-jacente et l’affichage mis en forme dans l’interface utilisateur Excel peut se traduire par une confusion et des erreurs pendant les calculs du complément. Les types de données personnalisés sont une solution à cet écart.

Les types de données développent la prise en charge de l’API JavaScript au-delà des quatre types de données d’origine (chaîne, nombre, booléen et erreur) pour inclure des images web, des valeurs numériques mises en forme, des valeurs d’entité, des tableaux au sein des valeurs d’entité et des types de données d’erreur améliorés en tant que structures de données flexibles. Ces types, qui permettent de nombreuses expériences de [types de données liées](https://support.microsoft.com/office/what-linked-data-types-are-available-in-excel-6510ab58-52f6-4368-ba0f-6a76c0190772), offrent une précision et une simplicité lors des calculs du complément et étendent le potentiel des compléments Excel au-delà d’une grille à 2 dimensions.

## <a name="data-types-and-custom-functions"></a>Types de données et fonctions personnalisées

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

Les types de données améliorent la puissance des fonctions personnalisées. Les fonctions personnalisées acceptent les types de données comme entrées et sorties de fonctions personnalisées et les fonctions personnalisées utilisent le même schéma JSON pour les types de données que l’API JavaScript Excel. Ce schéma JSON de types de données est conservé à mesure que les fonctions personnalisées calculent et évaluent. Si vous souhaitez en savoir plus sur l’intégration des types de données à vos fonctions personnalisées, consultez les [Concepts de base des fonctions personnalisées et des types de données](custom-functions-data-types-concepts.md).

## <a name="see-also"></a>Voir aussi

- [Concepts de base des types de données Excel](excel-data-types-concepts.md)
- [Référence de l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Vue d’ensemble des fonctions personnalisées et des types de données](custom-functions-data-types-overview.md)