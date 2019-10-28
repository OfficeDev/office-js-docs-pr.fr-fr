---
title: Concepts avancés de programmation avec l’API JavaScript Excel
description: ''
ms.date: 07/17/2019
localization_priority: Priority
ms.openlocfilehash: 8755b479543d48fcbbbf2bfa1ea93fb40af87ecf
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/24/2019
ms.locfileid: "37681927"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a>Concepts avancés de programmation avec l’API JavaScript Excel

Cet article s’appuie sur les informations contenues dans la rubrique [Concepts de base pour la programmation de l’API JavaScript Excel](excel-add-ins-core-concepts.md) pour décrire certains concepts plus avancés qui sont indispensables à la création de compléments complexes pour Excel 2016 ou version ultérieure.

## <a name="officejs-apis-for-excel"></a>API Office.js pour Excel

Un complément Excel interagit avec des objets dans Excel à l’aide de l’API JavaScript pour Office, qui inclut deux modèles d’objets JavaScript :

* **API JavaScript pour Excel** : inclut dans Office 2016, l’[API JavaScript Excel](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) fournit des objets fortement typés que vous pouvez utiliser pour accéder à des feuilles de calcul, des plages, des tableaux, des graphiques et bien plus encore.

* **API communes** : incluses dans Office 2013, les [API communes](/javascript/api/office) peuvent être utilisées pour accéder à des fonctionnalités telles qu’une interface utilisateur, des boîtes de dialogue et des paramètres du client, qui sont communes à plusieurs types d’applications Office.

Vous utiliserez probablement l’API JavaScript Excel pour développer la majorité des fonctionnalités des compléments destinés à Excel 2016 ou version ultérieure, vous utiliserez également des objets dans l’API commune. Par exemple :

- [Context](/javascript/api/office/office.context) : l’objet **Context** représente l’environnement d’exécution du complément et permet d’accéder à des objets clés de l’API. Il se compose de détails sur la configuration du classeur comme `contentLanguage` et `officeTheme`, et fournit des informations sur l’environnement d’exécution du complément comme `host` et `platform`. En outre, il fournit la méthode `requirements.isSetSupported()` que vous pouvez utiliser pour vérifier si l’ensemble de conditions requises spécifié est pris en charge par l’application Excel dans laquelle le complément est exécuté.

- [Document](/javascript/api/office/office.document) : L’objet **Document** fournit la méthode `getFileAsync()` que vous pouvez utiliser pour télécharger le fichier Excel dans lequel le complément est exécuté.

## <a name="requirement-sets"></a>Ensembles de conditions requises

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Le complément Office peut effectuer une vérification à l’exécution ou utiliser des ensembles de conditions requises spécifiés dans le manifeste pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour identifier les ensembles de conditions requises spécifiques disponibles sur chaque plateforme prise en charge, reportez-vous à [Ensembles de conditions requises de l’API JavaScript pour Excel](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets).

### <a name="checking-for-requirement-set-support-at-runtime"></a>Vérification de la prise en charge de l’ensemble de conditions requises à l’exécution

L’exemple de code suivant montre comment déterminer si l’application hôte dans laquelle le complément est en cours d’exécution prend en charge l’ensemble spécifié de conditions requises pour l’API.

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>Définition de la prise en charge de l’ensemble de conditions requises dans le manifeste

Vous pouvez utiliser l’[élément Requirements](/office/dev/add-ins/reference/manifest/requirements) dans le manifeste de complément pour spécifier les ensembles de conditions requises minimales et/ou les méthodes d’API que votre complément doit activer. Si la plateforme ou l’hôte Office ne prend pas en charge les ensembles de conditions requises ou les méthodes d’API spécifiées dans l’élément **Requirements** du manifeste, le complément ne s’exécute pas dans cet hôte ou cette plateforme et ne s’affiche pas dans des compléments dans **Mes compléments**.

L’exemple de code suivant montre l’élément **Requirements** dans un manifeste indiquant que le complément doit être chargé dans toutes les applications hôtes Office prenant en charge l’ensemble de conditions requises ExcelApi version 1.3 ou ultérieure.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> Pour rendre votre complément disponible sur toutes les plateformes d’un hôte Office, comme Excel sur le web, Windows et iPad, nous vous recommandons de vérifier la prise en charge des conditions requises lors de l’exécution au lieu de définir la prise en charge d’ensemble de conditions requises dans le manifeste.

### <a name="requirement-sets-for-the-officejs-common-api"></a>Ensembles de conditions requises pour l’API commune Office.js

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).

## <a name="loading-the-properties-of-an-object"></a>Chargement des propriétés d’un objet

Tout appel de la méthode `load()` sur un objet JavaScript pour Excel demande à l’API de charger l’objet dans la mémoire JavaScript lorsque la méthode `sync()` est exécutée. La méthode `load()` accepte une chaîne qui contient des noms délimités par des virgules de propriétés à charger ou un objet spécifiant des propriétés à charger, des options de pagination, etc.

> [!NOTE]
> si vous appelez la méthode `load()` sur un objet (ou collection) sans spécifier de paramètre, toutes les propriétés scalaires de l’objet (ou toutes les propriétés scalaires de tous les objets de la collection) sont chargées. Pour réduire la quantité de données transférées entre l’application hôte Excel et le complément, évitez d’appeler la méthode `load()` sans spécifier explicitement les propriétés à charger.

### <a name="method-details"></a>Détails de méthodes

#### <a name="loadparam-object"></a>load(param: object)

Remplit l’objet de proxy créé dans le calque JavaScript avec les valeurs de propriété et d’objet spécifiées par les paramètres.

#### <a name="syntax"></a>Syntaxe

```js
object.load(param);
```

#### <a name="parameters"></a>Paramètres

|**Paramètre**|**Type**|**Description**|
|:------------|:-------|:----------|
|`param`|objet|Facultatif. Accepte des noms de propriétés sous forme de tableau ou de chaîne délimitée par des virgules. Un objet peut également être transmis pour définir les propriétés de sélection et de navigation (comme illustré dans l’exemple ci-dessous).|

#### <a name="returns"></a>Retourne

void

#### <a name="example"></a>Exemple

L’exemple de code suivant définit les propriétés d’une plage Excel en copiant les propriétés d’une autre plage. L’objet source doit d’abord être chargé, avant que ses valeurs de propriété puissent être accessibles et écrites sur la plage cible. L’exemple suppose que les deux plages (**B2:E2** et **B7:E7**) comprennent des données, et que leur mise en forme initiale est différente.

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
|`select`|objet|Contient une liste délimitée par des virgules ou un tableau de propriétés scalaires. Facultatif.|
|`expand`|objet|Contient une liste délimitée par des virgules ou un tableau de propriétés de navigation. Facultatif.|
|`top`|int| Spécifie le nombre maximal d’éléments de collection qui peuvent être inclus dans le résultat. Facultatif. Vous pouvez utiliser cette option uniquement lorsque vous utilisez l’option de notation d’objet.|
|`skip`|int|Indiquez le nombre d’éléments de la collection devant être ignorés et exclus du résultat. Si une valeur est définie pour `top`, la sélection du jeu de résultats démarre une fois que le nombre spécifié d’éléments a été ignoré. Facultatif. Vous pouvez utiliser cette option uniquement lorsque vous utilisez l’option de notation d’objet.|

L’exemple de code suivant charge une collection de feuilles de calcul en sélectionnant la propriété `name` et l’élément `address` de la plage utilisée pour chaque feuille de calcul dans la collection. Il indique également que seules les cinq premières feuilles de calcul de la collection doivent être chargées. Vous pouvez traiter l’ensemble suivant de cinq feuilles de calcul en spécifiant `top: 10` et `skip: 5` en tant que valeurs d’attribut.

```js
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a>Propriétés scalaires et de navigation

Il existe deux catégories de propriétés: **scalaire** et **de navigation**. Les propriétés scalaires peuvent se voir attribuer des types, tels que des chaînes, des nombres entiers et des structures JSON. Les propriétés de navigation sont des objets en lecture seule et des collections d’objets dont les champs sont assignés, et non pas la propriété directement. Par exemple, les membres `name` et `position` sur l’objet [Worksheet](/javascript/api/excel/excel.worksheet) sont des propriétés scalaires, tandis que `protection` et `tables` sont des propriétés de navigation. `prompt` sur l’objet [DataValidation](/javascript/api/excel/excel.datavalidation) est un exemple de propriété scalaire qui doit être définie à l’aide d’un objet JSON (`dv.prompt = { title: "MyPrompt"}`), au lieu de définir les sous-propriétés (`dv.prompt.title = "MyPrompt" // will not set the title`).

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a>Propriétés scalaires et propriétés de navigation avec `object.load()`

Tout appel de la méthode `object.load()` sans paramètre spécifié charge toutes les propriétés scalaires de l’objet. Les propriétés de navigation de l’objet ne sont pas chargées. En outre, les propriétés de navigation ne peuvent pas être chargées directement. Au lieu de cela, vous devez utiliser la méthode `load()` pour référencer des propriétés scalaires individuelles au sein de la propriété de navigation de votre choix. Par exemple, pour charger le nom de la police d’une plage, vous devez spécifier les propriétés de navigation **format** et **font** en tant que chemin d’accès à la propriété **name** :

```js
someRange.load("format/font/name")
```

> [!NOTE]
> grâce à l’API JavaScript pour Excel, vous pouvez définir les propriétés scalaires d’une propriété de navigation en parcourant le chemin d’accès. Par exemple, vous pouvez définir la taille de police pour une plage à l’aide de `someRange.format.font.size = 10;`. Il est inutile de charger la propriété avant de la configurer. 

## <a name="setting-properties-of-an-object"></a>Définition des propriétés d’un objet

La définition de propriétés sur un objet avec des propriétés de navigation imbriquées peut être laborieuse. Au lieu de définir des propriétés individuelles à l’aide de chemins de navigation comme décrit ci-dessus, vous pouvez utiliser la méthode `object.set()` qui est disponible sur tous les objets de l’API JavaScript pour Excel. Grâce à cette méthode, vous pouvez définir plusieurs propriétés d’un objet à la fois en transmettant soit un autre objet du même type Office.js, soit un objet JavaScript avec des propriétés structurées comme celles de l’objet sur lequel la méthode est appelée.

> [!NOTE]
> la méthode `set()` est implémentée uniquement pour les objets dans les API JavaScript pour Office propres à un hôte, telles que l’API JavaScript pour Excel. Les API communes (partagées) ne prennent pas en charge cette méthode. 

### <a name="set-properties-object-options-object"></a>set (properties: object, options: object)

Les propriétés de l’objet sur lequel la méthode est appelée sont définies sur les valeurs spécifiées par les propriétés correspondantes de l’objet transmis. Si le paramètre `properties` est un objet JavaScript, toute propriété de l’objet transmis qui correspond à une propriété en lecture seule dans l’objet sur lequel la méthode est appelée sera ignorée ou générera une exception, en fonction de la valeur du paramètre `options`.

#### <a name="syntax"></a>Syntaxe

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a>Paramètres

|**Paramètre**|**Type**|**Description**|
|:------------|:--------|:----------|
|`properties`|objet|Objet de même type Office.js que l’objet sur lequel la méthode est appelée ou objet JavaScript avec des noms et des types de propriétés reflétant la structure de l’objet sur lequel la méthode est appelée.|
|`options`|objet|Facultatif. Peut être transmis uniquement si le premier paramètre est un objet JavaScript. L’objet peut contenir la propriété suivante : `throwOnReadOnly?: boolean` (La valeur par défaut est `true` : générer une erreur si l’objet JavaScript transmis inclut des propriétés en lecture seule.)|

#### <a name="returns"></a>Retourne

void

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

## <a name="42ornullobject-methods"></a>Méthodes &#42;OrNullObject

De nombreuses méthodes d’API JavaScript pour Excel renvoient une exception lorsque la condition de l’API n’est pas remplie. Par exemple, si vous tentez d’obtenir une feuille de calcul en spécifiant le nom d’une feuille de calcul qui n’existe pas dans le classeur, la méthode `getItem()` renvoie une exception `ItemNotFound`. 

Au lieu d’implémenter une logique complexe de gestion des exceptions pour des scénarios similaires, vous pouvez utiliser la variante de la méthode `*OrNullObject` disponible pour les différentes méthodes de l’API JavaScript pour Excel. Une méthode `*OrNullObject` renvoie un objet Null (pas l’élément JavaScript `null`) au lieu de lever une exception si l’élément spécifié n’existe pas. Par exemple, vous pouvez appeler la méthode `getItemOrNullObject()` sur une collection telle que **Worksheets** pour tenter de récupérer un élément de la collection. La méthode `getItemOrNullObject()` renvoie l’élément spécifié, s’il existe. Sinon, elle renvoie un objet Null. L’objet Null renvoyé contient la propriété booléenne `isNullObject` que vous pouvez étudier pour déterminer l’existence de l’objet.

L’exemple de code suivant tente de récupérer une feuille de calcul nommée « Data » à l’aide de la méthode `getItemOrNullObject()`. Si la méthode renvoie un objet Null, une nouvelle feuille doit être créée pour pouvoir réaliser des actions sur la feuille.

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
* [Optimisation des performances à l’aide de l’API JavaScript d’Excel](performance.md)
* [Référence de l’API JavaScript pour Excel](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
