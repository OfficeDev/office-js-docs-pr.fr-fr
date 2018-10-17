---
title: Concepts avancés de programmation avec l’API JavaScript Excel
description: ''
ms.date: 10/03/2018
ms.openlocfilehash: 09f2d95e4cf7631b519f00cddee265dbf697e07e
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505887"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a>Concepts avancés de programmation avec l’API JavaScript Excel

Cet article s’appuie sur les informations contenues dans la rubrique [Concepts de programmation fondamentaux de l’API JavaScript d’Excel](excel-add-ins-core-concepts.md) pour décrire certains des concepts les plus avancés qui sont essentiels à la création de compléments complexes pour Excel 2016 ou version ultérieure.

## <a name="officejs-apis-for-excel"></a>API Office.js pour Excel

Un complément Excel interagit avec des objets dans Excel à l’aide de l’API JavaScript pour Office, qui inclut deux modèles d’objets JavaScript :

* **API JavaScript pour Excel** : inclut dans Office 2016, l’[API JavaScript Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) fournit des objets fortement typés que vous pouvez utiliser pour accéder à des feuilles de calcul, des plages, des tableaux, des graphiques et bien plus encore. 

* **API communes** : incluses dans Office 2013, les API communes (également appelées [API partagées](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js)) peuvent être utilisées pour accéder à des fonctionnalités, telles que l’interface utilisateur, les boîtes de dialogue et les paramètres du client, qui sont communes à plusieurs types d’applications hôtes, comme Word, Excel et PowerPoint.

Bien que vous utiliserez probablement l'API JavaScript d'Excel pour développer la majorité des fonctionnalités dans des compléments destinés à Excel 2016 ou une version ultérieure, vous utiliserez également des objets dans l’API partagée. Par exemple : Par exemple :

- [Context](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js) : l’objet **Context** représente l’environnement d’exécution du complément et donne accès à des objets clés de l’API. Il se compose des détails de configuration de classeur, tels que `contentLanguage` et `officeTheme` et fournit des informations sur l’environnement d’exécution du complément comme `host` et `platform`. En outre, il fournit la méthode `requirements.isSetSupported()`, que vous pouvez utiliser pour vérifier si l’ensemble de conditions requises spécifié est pris en charge par l’application Excel où le complément est en cours d’exécution. 

- [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) : L’objet **Document** fournit la méthode `getFileAsync()` que vous pouvez utiliser pour télécharger le fichier Excel dans lequel le complément est exécuté. 

## <a name="requirement-sets"></a>Ensembles de conditions requises

Les ensembles de conditions requises sont des groupes nommés de membres de l’API. Un complément Office peut effectuer une vérification à l’exécution ou utiliser des ensembles de conditions requises spécifiés dans le manifeste pour déterminer si un hôte Office prend en charge les API dont le complément a besoin. Pour identifier les ensembles de conditions requises spécifiques qui sont disponibles sur chaque plate-forme prise en charge, voir [Ensembles de conditions requises de l’API JavaScript d’Excel](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js).

### <a name="checking-for-requirement-set-support-at-runtime"></a>Vérification de la prise en charge de l’ensemble de conditions requises à l’exécution

L’exemple de code suivant montre comment déterminer si l’application hôte dans laquelle le complément est en cours d’exécution prend en charge l’ensemble spécifié de conditions requises pour l’API.

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>Définition de la prise en charge de l’ensemble de conditions requises dans le manifeste

Vous pouvez utiliser l’[élément Requirements](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/requirements?view=office-js) dans le manifeste du complément pour spécifier les ensembles de conditions requises minimaux et/ou les méthodes API que votre complément doit activer. Si la plateforme ou l’hôte Office ne prend pas en charge les ensembles de conditions requises ou les méthodes API qui sont spécifiés dans l’élément **Requirements** du manifeste, le complément ne sera pas exécuté dans cet hôte ou cette plateforme et ne s’affichera pas dans la liste des compléments répertoriés dans **Mes Compléments**. 

L’exemple de code suivant montre l’élément **Requirements** dans un manifeste indiquant que le complément doit être chargé dans toutes les applications hôtes Office prenant en charge l’ensemble de conditions requises ExcelApi version 1.3 ou ultérieure.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> pour rendre votre complément disponible sur toutes les plateformes d’un hôte Office, comme Excel pour Windows, Excel Online et Excel pour iPad, nous vous recommandons de vérifier la prise en charge des conditions requises lors de l’exécution au lieu de définir la prise en charge d’ensemble de conditions requises dans le manifeste.

### <a name="requirement-sets-for-the-officejs-common-api"></a>Ensembles de conditions requises pour l’API commune Office.js

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js).

## <a name="loading-the-properties-of-an-object"></a>Chargement des propriétés d’un objet

Appeler la méthode `load()` sur un objet JavaScript Excel indique à l’API de charger l’objet en mémoire JavaScript lors de l’exécution de la méthode `sync()`. La méthode `load()` accepte une chaîne qui contient les noms de propriétés délimités par des virgules à charger ou un objet qui spécifie les propriétés à charger, les options de la pagination, etc. 

> [!NOTE]
> Si vous appelez la méthode `load()` sur un objet (ou une collection) sans spécifier aucun paramètre, toutes les propriétés scalaires de l’objet (ou toutes les propriétés scalaires de tous les objets de la collection) seront chargées. Pour réduire la quantité de transfert de données entre l’application hôte Excel et le complément, vous devez éviter d’appeler la méthode `load()` sans spécifier explicitement les propriétés à charger.

### <a name="method-details"></a>Détails de méthodes

#### <a name="loadparam-object"></a>load(param: object)

Remplit l’objet proxy créé dans la couche JavaScript avec les valeurs de propriété et d’objet spécifiées par les paramètres.

#### <a name="syntax"></a>Syntaxe

```js
object.load(param);
```

#### <a name="parameters"></a>Paramètres

|**Paramètre**|**Type**|**Description**|
|:------------|:-------|:----------|
|`param`|objet|Facultatif. Accepte les noms de paramètre et de relation en tant que chaîne délimitée par des virgules ou que tableau. Un objet peut également être passé pour définir les propriétés de sélection et de navigation (comme illustré dans l’exemple ci-dessous).|

#### <a name="returns"></a>Renvoie

annuler

#### <a name="example"></a>Exemple

L’exemple de code suivant définit les propriétés d’une plage Excel en copiant les propriétés d’une autre plage. Notez que l’objet source doit être chargé en premier, avant que ses valeurs de propriété soient accessibles et écrites dans la plage cible. Cet exemple suppose qu’il existe des données dans deux plages (**B2:E2** et **B7:E7**) et que les deux plages sont initialement formatées différemment.

```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var sourceRange = sheet.getRange("B2:E2");
    sourceRange.load("format/fill/color, format/font/name, format/font/color");

    return ctx.sync()
        .then(function () {
            var targetRange = sheet.getRange("B7:E7");
            targetRange.set(sourceRange); 
            targetRange.format.autofitColumns();

            return ctx.sync();
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="load-option-properties"></a>Charger des propriétés d’option

Au lieu de transmettre un tableau ou une chaîne délimitée par des virgules lorsque vous appelez la méthode `load()`, vous pouvez également transmettre un objet qui contient les propriétés suivantes. 

|**Propriété**|**Type**|**Description**|
|:-----------|:-------|:----------|
|`select`|objet|Contient une liste délimitée par des virgules ou un tableau de noms de paramètres/relations. Facultatif.|
|`expand`|objet|Contient une liste délimitée par des virgules ou un tableau de noms de relations. Facultatif.|
|`top`|int| Spécifie le nombre maximal d’éléments de collection qui peuvent être inclus dans le résultat. Facultatif. Vous pouvez utiliser cette option uniquement lorsque vous utilisez l’option de notation d’objet.|
|`skip`|int|Indiquez le nombre d’éléments de la collection devant être ignorés et exclus du résultat. Si une valeur est définie pour `top`, la sélection du jeu de résultats démarre une fois que le nombre spécifié d’éléments a été ignoré. Facultatif. Vous pouvez utiliser cette option uniquement lorsque vous utilisez l’option de notation d’objet.|

L’exemple de code suivant charge une collection de feuilles de calcul en sélectionnant la propriété  `name` et le `address` de la plage utilisée pour chaque feuille de calcul de la collection. Il spécifie également que seules les cinq premières feuilles de calcul de la collection doivent être chargées. Vous pouvez traiter l’ensemble de cinq feuilles de calcul suivant en spécifiant `top: 10` et `skip: 5` en tant que valeurs d’attribut. 

```js 
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a>Propriétés scalaires et de navigation 

Dans la documentation de référence de l’API JavaScript d’Excel, vous pouvez remarquer que les membres d’objet sont regroupés en deux catégories : les **propriétés** et les **relations**. Une propriété d’un objet est un membre scalaire tel qu’une chaîne, un entier ou une valeur booléenne, alors qu’une relation de objet (également appelée propriété de navigation) est un membre qui est soit un objet ou une collection d’objets. Par exemple, les membres `name` et `position` de l’objet [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=office-js) sont des propriétés scalaires, tandis que `protection` et `tables` sont des relations (propriétés de navigation). 

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a>Propriétés scalaires et propriétés de navigation avec `object.load()`

Appeler la méthode `object.load()` sans aucun paramètre spécifié charge toutes les propriétés scalaires de l’objet ; les propriétés de navigation de l’objet ne seront pas chargées. En outre, les propriétés de navigation ne peuvent pas être chargées directement. Au lieu de cela, vous devez utiliser la méthode `load()` pour référencer les propriétés scalaires individuelles au sein de la propriété de navigation de votre choix. Par exemple, pour charger le nom de la police pour une plage, vous devez spécifier les propriétés de navigation **format** et **font** en tant que chemin d’accès à la propriété **name** :

```js
someRange.load("format/font/name")
```

> [!NOTE]
> Avec l’API JavaScript d’Excel, vous pouvez définir les propriétés scalaires d’une propriété de navigation en parcourant le chemin d’accès. Par exemple, vous pourriez définir la taille de police pour une plage à l’aide de `someRange.format.font.size = 10;`. Il est inutile de charger la propriété avant de la définir. 

## <a name="setting-properties-of-an-object"></a>Définition des propriétés d’un objet

La définition des propriétés sur un objet avec des propriétés de navigation imbriquées peut être fastidieux. Au lieu de définir des propriétés individuelles à l’aide de chemins d’accès de navigation comme indiqué ci-dessus, vous pouvez utiliser la méthode `object.set()` disponible sur tous les objets dans l’interface API JavaScript d’Excel. Avec cette méthode, vous pouvez définir plusieurs propriétés d’un objet à la fois en transmettant un autre objet du même type Office.js ou un objet JavaScript avec des propriétés structurées comme les propriétés de l’objet sur lequel la méthode est appelée.

> [!NOTE]
> La méthode `set()` est implémentée uniquement pour les objets des APIs JavaScript Office spécifiques à l’hôte, telles que l’interface API JavaScript d’Excel. Les API communes (partagées) ne prennent pas charge cette méthode. 

### <a name="set-properties-object-options-object"></a>set (properties: object, options: object)

Les propriétés de l’objet sur lequel la méthode est appelée sont définies sur les valeurs spécifiées par les propriétés correspondantes de l’objet transmis. Si le paramètre `properties` est un objet JavaScript, toute propriété de l’objet transmis qui correspond à une propriété en lecture seule dans l’objet sur lequel la méthode est appelée sera ignorée ou générera une exception, en fonction de la valeur du paramètre `options`.

#### <a name="syntax"></a>Syntaxe

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a>Paramètres

|**Paramètre**|**Type**|**Description**|
|:------------|:--------|:----------|
|`properties`|objet|Objet de même type Office.js que l’objet sur lequel la méthode est appelée ou objet JavaScript avec des noms et des types de propriétés reflétant la structure de l’objet sur lequel la méthode est appelée.|
|`options`|objet|Facultatif. Peut être transmis uniquement si le premier paramètre est un objet JavaScript. L’objet peut contenir la propriété suivante : `throwOnReadOnly?: boolean` (La valeur par défaut est `true` : générer une erreur si l’objet JavaScript transmis inclut des propriétés en lecture seule.)|

#### <a name="returns"></a>Renvoie

annuler    

#### <a name="example"></a>Exemple

L’exemple de code suivant définit plusieurs propriétés de mise en forme d’une plage en appelant la méthode `set()` et en transmettant un objet JavaScript avec des noms et des types de propriétés reflétant la structure des propriétés dans l’objet **Range**. Cet exemple part du principe que des données sont présentes dans la plage **B2:E2**.

```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync(); 
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
## <a name="42ornullobject-methods"></a>Méthodes *OrNullObject

De nombreuses méthodes de l’interface API JavaScript d’Excel renvoient une exception lorsque la condition de l’API n’est pas remplie. Par exemple, si vous tentez d’obtenir une feuille de calcul en spécifiant un nom de feuille de calcul qui n’existe pas dans le classeur, la méthode `getItem()` renverra une exception `ItemNotFound`. 

Au lieu d’implémenter des logiques de gestion d’exceptions complexes pour des scénarios comme celui-ci, vous pouvez utiliser la variante de méthode `*OrNullObject` qui est disponible pour plusieurs méthodes dans l’API Javascript d’Excel. Une méthode `*OrNullObject` retournera un objet null (pas le `null` Javascript) au lieu de générer une exception si l’élément spécifié n’existe pas. Par exemple, vous pouvez appeler la méthode `getItemOrNullObject()` sur une collection comme **Worksheets** pour tenter de récupérer un élément dans une collection. La méthode `getItemOrNullObject()` renvoie l’élément spécifié s’il existe ; sinon, elle renvoie un objet null. L’objet null renvoyé contient la propriété booléenne `isNullObject` que vous pouvez évaluer pour déterminer si l’objet existe.

L’exemple de code suivant tente de récupérer une feuille de calcul nommée « Data » à l’aide de la méthode `getItemOrNullObject()`. Si la méthode renvoie un objet null, une nouvelle feuille doit être créée avant que des actions puissent être menées sur la feuille.

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data"); 

return context.sync()
  .then(function() {
    if (dataSheet.isNullObject) { 
        // Create the sheet
    }

    dataSheet.position = 1;
    //...
  })
```

## <a name="see-also"></a>Voir aussi
 
* [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
* [Exemples de code pour les compléments Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Optimisation des performances de l'API JavaScript d'Excel](performance.md)
* [Référence de l’API JavaScript pour Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
