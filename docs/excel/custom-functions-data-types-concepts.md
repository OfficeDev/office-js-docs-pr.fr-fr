---
title: Concepts de base des fonctions personnalisées et des types de données
description: Découvrez les concepts de base pour l’utilisation Excel types de données avec vos fonctions personnalisées.
ms.date: 11/03/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
ms.openlocfilehash: 3b7e735f78ca7b6dcdffa3bd5e8ba9c9d3093766
ms.sourcegitcommit: ad5d7ab21f64012543fb2bd9226d90330d25468b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/04/2021
ms.locfileid: "60749405"
---
# <a name="custom-functions-and-data-types-core-concepts-preview"></a>Fonctions personnalisées et concepts de base des types de données (aperçu)

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

Les types de données améliorent Excel l’API JavaScript en étendu la prise en charge des types de données au-delà des quatre types d’origine (chaîne, nombre, booléen et erreur). Les types de données incluent la prise en charge des valeurs numériques formatées, des images web, des valeurs d’entité et des tableaux au sein des valeurs d’entité. Les fonctions personnalisées acceptent les types de données en tant que valeurs d’entrée et de sortie, ce qui développe la puissance de calcul des fonctions personnalisées.

Pour en savoir plus sur l’utilisation des types de données avec un Excel, voir Excel concepts de [base des types de données.](excel-data-types-concepts.md)

## <a name="how-custom-functions-handle-data-types"></a>Gestion des types de données par les fonctions personnalisées

Les fonctions personnalisées peuvent reconnaître les types de données et les accepter comme valeurs de paramètre. Une fonction personnalisée peut créer un nouveau type de données pour une valeur de retour. Les fonctions personnalisées utilisent le même schéma JSON pour les types de données que l’API JavaScript Excel, et ce schéma JSON est conservé à mesure que les fonctions personnalisées calculent et évaluent.

> [!NOTE]
> Les fonctions personnalisées ne prisent pas en charge toutes les fonctionnalités des objets d’erreur améliorés offerts par les types de données. Une fonction personnalisée peut accepter un objet d’erreur de types de données, mais elle ne sera pas conservée tout au long du calcul. Pour l’instant, les fonctions personnalisées ne peuvent que supporter les [erreurs incluses dans l’objet CustomFunctions.Error.](custom-functions-errors.md)

## <a name="enable-data-types-for-custom-functions"></a>Activer les types de données pour les fonctions personnalisées

Pour utiliser cette fonctionnalité, vous devez mettre à jour manuellement vos métadonnées JSON. Pour des tests plus temporaires, vous pouvez personnaliser vos paramètres de Script Lab au lieu de mettre à jour manuellement les métadonnées JSON. Les sections suivantes décrivent ces étapes plus en détail.

### <a name="manually-update-json-metadata"></a>Mettre à jour manuellement les métadonnées JSON

Les projets de fonctions personnalisées incluent un fichier de métadonnées JSON. Ce fichier de métadonnées JSON diffère du schéma JSON utilisé par les API de types de données. Pour utiliser l’intégration des types de données avec des fonctions personnalisées, le fichier de métadonnées JSON des fonctions personnalisées doit être mis à jour manuellement pour inclure la propriété `allowCustomDataForDataTypeAny` . Définissez cette propriété sur `true` .

Pour obtenir une description complète du processus de création manuelle JSON, voir Créer manuellement des métadonnées [JSON pour les fonctions personnalisées.](custom-functions-json.md) Pour plus d’informations sur cette propriété, [voir allowCustomDataForDataTypeAny.](custom-functions-json.md#allowcustomdatafordatatypeany-preview)

### <a name="script-lab-option"></a>Script Lab option

L’intégration des fonctions personnalisées avec les types de données est disponible pour les tests avec Script Lab, en plus de la mise à jour manuelle des métadonnées JSON décrite dans la section précédente. Pour en savoir plus sur Script Lab, voir [Explorer Office API JavaScript à l’aide Script Lab](../overview/explore-with-script-lab.md). Pour tester cette fonctionnalité avec Script Lab, mettez à jour les paramètres en suivant les étapes ci-après.

1. Ouvrez le Script Lab **de tâches Code.**
1. Dans le coin inférieur droit, **sélectionnez** Paramètres bouton.
1. Go to the **User Paramètres** tab and enter `allowCustomDataForDataTypeAny: true` .

![Capture d’écran montrant les étapes à suivre pour activer les types de données pour les fonctions personnalisées dans Script Lab.](../images/custom-functions-script-lab-data-type.png)

## <a name="output-a-formatted-number-value"></a>Sortie d’une valeur numérique mise en forme

L’exemple de code suivant montre comment créer un type de données [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) avec une fonction personnalisée. La fonction prend un nombre de base et un paramètre de format comme paramètres d’entrée et renvoie un type de données de valeur numérique mise en forme en tant que sortie.

```js
/**
 * Take a number as the input value and return a formatted number value as the output.
 * @customfunction
 * @param {number} value
 * @param {string} format (e.g. "0.00%")
 * @returns A formatted number value.
 */
function createFormattedNumber(value, format) {
    return {
        type: "FormattedNumber",
        basicValue: value,
        numberFormat: format
    }
}
```

## <a name="input-an-entity-value"></a>Saisie d’une valeur d’entité

L’exemple de code suivant montre une fonction personnalisée qui prend un type de données [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) comme entrée. Si le `attribute` paramètre est définie sur , la fonction renvoie la propriété de la valeur `text` `text` d’entité. Sinon, la fonction renvoie la `basicValue` propriété de la valeur d’entité.

```js
/**
 * Accept an entity value data type as a function input.
 * @customfunction
 * @param {any} value
 * @param {string} attribute
 * @returns {any} The text value of the entity.
 */
function getEntityAttribute(value, attribute) {
    if (value.type == "Entity") {
        if (attribute == "text") {
            return value.text;
        } else {
            return value.properties[attribute].basicValue;
        }
    } else {
        return JSON.stringify(value);
    }
}
```

## <a name="see-also"></a>Voir aussi

* [Vue d’ensemble des fonctions personnalisées et des types de données](custom-functions-data-types-overview.md)
* [Présentation des types de données dans les compléments Excel](excel-data-types-overview.md)
* [Concepts de base des types de données Excel](excel-data-types-concepts.md)
* [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
