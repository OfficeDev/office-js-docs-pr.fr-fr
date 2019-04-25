---
ms.date: 03/20/2019
description: Découvrez comment exécuter les fonctions personnalisées dans Excel.
title: Architecture de fonctions personnalisées (aperçu)
localization_priority: Priority
ms.openlocfilehash: b3f3d6c5eda51639a734c6d0f162c596f0c1e41b
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448601"
---
# <a name="custom-functions-architecture"></a>Architecture de fonctions

 Les fonctions personnalisées sont avec leur propre runtime unique qui privilégie l’exécution de calculs. Cet article traitera les différences entre le runtime de fonctions personnalisées et le moteur JavaScript basés sur navigateur qui alimente la plupart des autres parties de votre complément.

## <a name="custom-functions-runtime"></a>Runtime des fonctions personnalisées

Un complément Web Office peut interagir avec l’utilisateur en tant que volet Office ou volet de contenu et peut inclure des commandes et fonctions personnalisées. Toutes ces parties s’exécutent dans un moteur de navigateur d’exécution à l’exception des fonctions personnalisées. Les fonctions personnalisées s’exécutent dans un runtime de fonctions personnalisées distincte pour optimiser la vitesse de calcul.

Notez que si vous utilisez le [Générateur Yeoman de compléments Office](https://www.npmjs.com/package/generator-office) pour générer votre projet, le runtime de fonctions personnalisées est chargé via le fichier de script personnalisé fonctions.js référencé dans le fichier fonctions.html. La fonctions.html sert uniquement à charger le runtime et ne doit pas être utilisé comme le volet Office pour votre complément.

Le tableau suivant met en évidence les différences entre l’exécution de fonctions personnalisées et l’exécution du moteur navigateur:

| Exécution des fonctions personnalisées  | Exécution du moteur navigateur    |
|------------------------------------------------------------------ |-------------------------------------------------------------------------------------------------------------- |
| Prend en charge le renvoi d’une valeur d’une cellule    | Prend en charge les éléments Office.js APIs et éléments d’Interface Utilisateurs   |
| N’a pas d’`localStorage`objet, à la place utilise`AsyncStorage`  | A`localStorage` d’objet, peut éventuellement utiliser l’`AsyncStorage`objet   |
| Ne prend pas en charge interaction avec le DOM ou le chargement des bibliothèques qui dépendent de DOM par exemple, jQuery.    | Ne prend pas en charge l’interaction avec le DOM ou le chargement des bibliothèques qui dépendent de DOM. |


## <a name="browser-engine-runtime"></a>Exécution du moteur navigateur

Le volet de tâche, complément de contenu et les commandes s’exécutent dans une navigateur d’exécution du moteur.

L’exécution du moteur navigateur prend en charge les APIs Office.js. N’oubliez pas que les API Excel, telles que des API qui vous permettent de manipuler des tableaux Excel, exécutent sur le runtime moteur de navigateur, mais ne sont pas accessibles directement à partir de l’exécution de fonctions personnalisées.

## <a name="communicate-between-runtimes"></a>Communiquer entre les exécutions

Votre code de fonctions personnalisées ne peut pas interagir directement avec le code d’autres parties de votre complément web, par exemple, le volet de tâche, car elles se trouvent dans différentes exécutions. Mais dans certains cas, vous devrez partager des données, par exemple, en passant un jeton.

`AsyncStorage` permet de stocker des données à partir de vos fonctions personnalisées et obtenir des données à partir de votre code de volet de tâches. Pour plus d’informations sur le stockage et partage de données, voir [Enregistrer et partager état](custom-functions-overview.md#saving-and-sharing-state).

Vous pouvez voir un exemple de code à l’aide `AsyncStorage` dans cet [référentiel Github](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) dédiée aux pratiques et modèles.
Pour des informations plus générales sur `AsyncStorage`, voir [Runtime de Fonctions personnalisées](./custom-functions-runtime.md).

`AsyncStorage` peut également être utile pour l’authentification. Pour plus d’informations, voir[Authentification des fonctions personnalisées](custom-functions-authentication.md).

## <a name="see-also"></a>Voir aussi

* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Fonctions personnalisées changelog](custom-functions-changelog.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
