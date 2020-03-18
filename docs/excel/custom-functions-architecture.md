---
ms.date: 07/10/2019
description: En savoir plus sur le runtime pour les fonctions personnalisées Excel.
title: Architecture de fonctions personnalisées
localization_priority: Normal
ms.openlocfilehash: a11ac40591e11725bb35b16bf53fa07062541c8f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718768"
---
# <a name="custom-functions-architecture"></a>Architecture de fonctions

 Les fonctions personnalisées sont avec leur propre runtime unique qui privilégie l’exécution de calculs. Cet article traitera les différences entre le runtime de fonctions personnalisées et le moteur JavaScript basés sur navigateur qui alimente la plupart des autres parties de votre complément.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="custom-functions-runtime"></a>Runtime des fonctions personnalisées

Un complément Web Office peut interagir avec l’utilisateur en tant que volet Office ou volet de contenu et peut inclure des commandes et fonctions personnalisées. Toutes ces parties s’exécutent dans un moteur de navigateur d’exécution à l’exception des fonctions personnalisées. Les fonctions personnalisées s’exécutent dans un runtime de fonctions personnalisées distincte pour optimiser la vitesse de calcul.

Notez que si vous utilisez le [Générateur Yeoman de compléments Office](https://www.npmjs.com/package/generator-office) pour générer votre projet, le runtime de fonctions personnalisées est chargé via le fichier de script personnalisé fonctions.js référencé dans le fichier **fonctions.html**. La **fonctions.html** sert uniquement à charger le runtime et ne doit pas être utilisé comme le volet Office pour votre complément.

Le tableau suivant met en évidence les différences entre l’exécution de fonctions personnalisées et l’exécution du moteur navigateur:

| Exécution des fonctions personnalisées     | Exécution du moteur navigateur     |
|------------------------------------------------------------------    |--------------------------------------------------------------------------------------------------------------    |
| Prend en charge le renvoi d’une valeur d’une cellule     | Prend en charge les éléments Office.js APIs et éléments d’Interface Utilisateurs     |
| N’a pas l’objet `localStorage`, utilise à la place l’objet `OfficeRuntime.storage`.     | A l’objet `localStorage`, peut éventuellement utiliser l’objet `OfficeRuntime.storage`.     |
| Ne prend pas en charge interaction avec le DOM ou le chargement des bibliothèques qui dépendent de DOM par exemple, jQuery.    | Ne prend pas en charge l’interaction avec le DOM ou le chargement des bibliothèques qui dépendent de DOM. |

## <a name="browser-engine-runtime"></a>Exécution du moteur navigateur

Le volet de tâche, complément de contenu et les commandes s’exécutent dans une navigateur d’exécution du moteur.

L’exécution du moteur navigateur prend en charge les APIs Office.js. N’oubliez pas que les API Excel, telles que des API qui vous permettent de manipuler des tableaux Excel, exécutent sur le runtime moteur de navigateur, mais ne sont pas accessibles directement à partir de l’exécution de fonctions personnalisées.

## <a name="communicate-between-runtimes"></a>Communiquer entre les exécutions

Votre code de fonctions personnalisées ne peut pas interagir directement avec le code d’autres parties de votre complément web, par exemple, le volet de tâche, car elles se trouvent dans différentes exécutions. Mais dans certains cas, vous devrez partager des données, par exemple, en passant un jeton.

L’objet `OfficeRuntime.storage` permet de stocker des données à partir de vos fonctions personnalisées et d’obtenir des données à partir de votre code de volet de tâches. Pour plus d’informations sur le stockage et partage de données, voir [Enregistrer et partager l’état](custom-functions-save-state.md).

Vous pouvez voir un exemple de code à l’aide de l’objet `storage` dans ce [référentiel Github](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) dédié aux pratiques et modèles.
Pour des informations plus générales sur `storage`, voir [Runtime de Fonctions personnalisées](./custom-functions-runtime.md).

L’objet `storage` peut également être utile pour l’authentification. Pour plus d’informations, voir[Authentification des fonctions personnalisées](custom-functions-authentication.md).

## <a name="next-steps"></a>Étapes suivantes
En savoir plus sur l' [utilisation de runtime des fonctions personnalisées](custom-functions-runtime.md).

## <a name="see-also"></a>Voir aussi

* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Recevoir et gérer des données à l’aide de fonctions personnalisées](custom-functions-web-reqs.md)
* [Métadonnées de fonctions personnalisées](custom-functions-json.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
