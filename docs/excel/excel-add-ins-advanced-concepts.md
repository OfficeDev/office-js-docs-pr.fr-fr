---
title: Concepts avanc?s pour l?API JavaScript Excel
description: ''
ms.date: 1/18/2018
ms.openlocfilehash: 89db69e124475c882448a2105837787ce2c84753
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="excel-javascript-api-advanced-concepts"></a>Concepts avanc?s pour l?API JavaScript Excel

Cet article s?appuie sur les informations contenues dans la rubrique [Concepts de base de l?API JavaScript Excel](excel-add-ins-core-concepts.md) pour d?crire certains concepts plus avanc?s qui sont indispensables ? la cr?ation de compl?ments complexes pour Excel 2016. 

## <a name="officejs-apis-for-excel"></a>API Office.js pour Excel

Un compl?ment Excel interagit avec des objets dans Excel ? l?aide de l?API JavaScript pour Office, qui inclut deux mod?les d?objets JavaScript :

* **API JavaScript pour Excel** : inclut dans Office 2016, l?[API JavaScript Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) fournit des objets fortement typ?s que vous pouvez utiliser pour acc?der ? des feuilles de calcul, des plages, des tableaux, des graphiques et bien plus encore. 

* **API communes** : incluses dans Office 2013, les API communes (?galement appel?es [API partag?es](https://dev.office.com/reference/add-ins/javascript-api-for-office)) peuvent ?tre utilis?es pour acc?der ? des fonctionnalit?s, telles que l?interface utilisateur, les bo?tes de dialogue et les param?tres du client, qui sont communes ? plusieurs types d?applications h?tes, comme Word, Excel et PowerPoint.

Lorsque vous emploierez l?API JavaScript Excel pour d?velopper la majorit? des fonctionnalit?s dans des compl?ments destin?s ? Excel 2016, vous utiliserez ?galement des objets dans l?API partag?e. Par exemple :

- [Context](https://dev.office.com/reference/add-ins/shared/context) : l?objet **Context** repr?sente l?environnement d?ex?cution du compl?ment et permet d?acc?der ? des objets cl?s de l?API. Il se compose de d?tails sur la configuration du classeur comme `contentLanguage` et `officeTheme`, et fournit des informations sur l?environnement d?ex?cution du compl?ment comme `host` et `platform`. En outre, il fournit la m?thode `requirements.isSetSupported()` que vous pouvez utiliser pour v?rifier si l?ensemble de conditions requises sp?cifi? est pris en charge par l?application Excel dans laquelle le compl?ment est ex?cut?. 

- [Document](https://dev.office.com/reference/add-ins/shared/document) : L?objet **Document** fournit la m?thode `getFileAsync()` que vous pouvez utiliser pour t?l?charger le fichier Excel dans lequel le compl?ment est ex?cut?. 

## <a name="requirement-sets"></a>Ensembles de conditions requises

Les ensembles de conditions requises sont des groupes nomm?s de membres d?API. Le compl?ment Office peut effectuer une v?rification ? l?ex?cution ou utiliser des ensembles de conditions requises sp?cifi?s dans le manifeste pour d?terminer si un h?te Office prend en charge les API requises par le compl?ment. Pour identifier les ensembles de conditions requises sp?cifiques disponibles sur chaque plateforme prise en charge, reportez-vous ? [Ensembles de conditions requises de l?API JavaScript pour Excel](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets).

### <a name="checking-for-requirement-set-support-at-runtime"></a>V?rification de la prise en charge de l?ensemble de conditions requises ? l?ex?cution

L?exemple de code suivant montre comment d?terminer si l?application h?te dans laquelle le compl?ment est en cours d?ex?cution prend en charge l?ensemble sp?cifi? de conditions requises pour l?API.

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>D?finition de la prise en charge de l?ensemble de conditions requises dans le manifeste

Vous pouvez utiliser l?[?l?ment Requirements](https://dev.office.com/reference/add-ins/manifest/requirements) dans le manifeste de compl?ment pour sp?cifier les ensembles de conditions requises minimales et/ou les m?thodes d?API que votre compl?ment doit activer. Si la plateforme ou l?h?te Office ne prend pas en charge les ensembles de conditions requises ou les m?thodes d?API sp?cifi?es dans l??l?ment **Requirements** du manifeste, le compl?ment ne s?ex?cute pas dans cet h?te ou cette plateforme et ne s?affiche pas dans des compl?ments dans **Mes compl?ments**. 

L?exemple de code suivant montre l??l?ment **Requirements** dans un manifeste indiquant que le compl?ment doit ?tre charg? dans toutes les applications h?tes Office prenant en charge l?ensemble de conditions requises ExcelApi version 1.3 ou ult?rieure.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> pour rendre votre compl?ment disponible sur toutes les plateformes d?un h?te Office, comme Excel pour Windows, Excel Online et Excel pour iPad, nous vous recommandons de v?rifier la prise en charge des conditions requises lors de l?ex?cution au lieu de d?finir la prise en charge d?ensemble de conditions requises dans le manifeste.

### <a name="requirement-sets-for-the-officejs-common-api"></a>Ensembles de conditions requises pour l?API commune Office.js

Pour plus d?informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).

## <a name="loading-the-properties-of-an-object"></a>Chargement des propri?t?s d?un objet

Tout appel de la m?thode `load()` sur un objet JavaScript pour Excel demande ? l?API de charger l?objet dans la m?moire JavaScript lorsque la m?thode `sync()` est ex?cut?e. La m?thode `load()` accepte une cha?ne qui contient des noms d?limit?s par des virgules de propri?t?s ? charger ou un objet sp?cifiant des propri?t?s ? charger, des options de pagination, etc. 

> [!NOTE]
> si vous appelez la m?thode `load()` sur un objet (ou collection) sans sp?cifier de param?tre, toutes les propri?t?s scalaires de l?objet (ou toutes les propri?t?s scalaires de tous les objets de la collection) sont charg?es. Pour r?duire la quantit? de donn?es transf?r?es entre l?application h?te Excel et le compl?ment, ?vitez d?appeler la m?thode `load()` sans sp?cifier explicitement les propri?t?s ? charger.

### <a name="method-details"></a>D?tails de m?thodes

#### <a name="loadparam-object"></a>load(param: object)

Remplit l?objet de proxy cr?? dans le calque JavaScript avec les valeurs de propri?t? et d?objet sp?cifi?es par les param?tres.

#### <a name="syntax"></a>Syntaxe

```js
object.load(param);
```

#### <a name="parameters"></a>Param?tres

|**Param?tre**|**Type**|**Description**|
|:------------|:-------|:----------|
|`param`|objet|Facultatif. Accepte des noms de param?tre et de relation sous forme de tableau ou de cha?ne d?limit?e par des virgules. Un objet peut ?galement ?tre transmis pour d?finir les propri?t?s de s?lection et de navigation (comme illustr? dans l?exemple ci-dessous).|

#### <a name="returns"></a>Retourne

void

#### <a name="example"></a>Exemple

L?exemple de code suivant d?finit les propri?t?s d?une plage Excel en copiant les propri?t?s d?une autre plage. L?objet source doit d?abord ?tre charg?, avant que ses valeurs de propri?t? puissent ?tre accessibles et ?crites sur la plage cible. L?exemple suppose que les deux plages (**B2:E2** et **B7:E7**) comprennent des donn?es, et que leur mise en forme initiale est diff?rente.

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

### <a name="load-option-properties"></a>Charger des propri?t?s d?option

Au lieu de transmettre un tableau ou une cha?ne d?limit?e par des virgules lorsque vous appelez la m?thode `load()`, vous pouvez ?galement transmettre un objet qui contient les propri?t?s suivantes. 

|**Propri?t?**|**Type**|**Description**|
|:-----------|:-------|:----------|
|`select`|objet|Contient une liste d?limit?e par des virgules ou un tableau de noms de param?tres/relations. Facultatif.|
|`expand`|objet|Contient une liste d?limit?e par des virgules ou un tableau de noms de relations. Facultatif.|
|`top`|int| Sp?cifie le nombre maximal d??l?ments de collection qui peuvent ?tre inclus dans le r?sultat. Facultatif. Vous pouvez utiliser cette option uniquement lorsque vous utilisez l?option de notation d?objet.|
|`skip`|int|Indiquez le nombre d??l?ments de la collection devant ?tre ignor?s et exclus du r?sultat. Si une valeur est d?finie pour `top`, la s?lection du jeu de r?sultats d?marre une fois que le nombre sp?cifi? d??l?ments a ?t? ignor?. Facultatif. Vous pouvez utiliser cette option uniquement lorsque vous utilisez l?option de notation d?objet.|

L?exemple de code suivant charge une collection en s?lectionnant la propri?t? `name` et l??l?ment `address` de la plage utilis?e pour chaque feuille de calcul dans la collection. Il indique ?galement que seules les cinq premi?res feuilles de calcul de la collection doivent ?tre charg?es. Vous pouvez traiter l?ensemble suivant de cinq feuilles de calcul en sp?cifiant `top: 10` et `skip: 5` en tant que valeurs d?attribut. 

```js 
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a>Propri?t?s scalaires et de navigation 

Dans la documentation de r?f?rence de l?API JavaScript pour Excel, les membres de l?objet sont regroup?s en deux cat?gories : les **propri?t?s** et les **relations**. Une propri?t? d?objet est un membre scalaire comme une cha?ne, un nombre entier ou une valeur bool?enne, alors qu?une relation d?objet (?galement appel?e propri?t? de navigation) est un membre qui est un objet ou une collection d?objets. Par exemple, les membres `name` et `position` sur l?objet [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheet) sont des propri?t?s scalaires, tandis que `protection` et `tables` sont des relations (propri?t?s de navigation). 

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a>Propri?t?s scalaires et propri?t?s de navigation avec `object.load()`

Tout appel de la m?thode `object.load()` sans param?tre sp?cifi? charge toutes les propri?t?s scalaires de l?objet. Les propri?t?s de navigation de l?objet ne sont pas charg?es. En outre, les propri?t?s de navigation ne peuvent pas ?tre charg?es directement. Au lieu de cela, vous devez utiliser la m?thode `load()` pour r?f?rencer des propri?t?s scalaires individuelles au sein de la propri?t? de navigation de votre choix. Par exemple, pour charger le nom de la police d?une plage, vous devez sp?cifier les propri?t?s de navigation **format** et **font** en tant que chemin d?acc?s ? la propri?t? **name** :

```js
someRange.load("format/font/name")
```

> [!NOTE]
> gr?ce ? l?API JavaScript pour Excel, vous pouvez d?finir les propri?t?s scalaires d?une propri?t? de navigation en parcourant le chemin d?acc?s. Par exemple, vous pouvez d?finir la taille de police pour une plage ? l?aide de `someRange.format.font.size = 10;`. Il est inutile de charger la propri?t? avant de la configurer. 

## <a name="setting-properties-of-an-object"></a>D?finition des propri?t?s d?un objet

La d?finition de propri?t?s sur un objet avec des propri?t?s de navigation imbriqu?es peut ?tre laborieuse. Au lieu de d?finir des propri?t?s individuelles ? l?aide de chemins de navigation comme d?crit ci-dessus, vous pouvez utiliser la m?thode `object.set()` qui est disponible sur tous les objets de l?API JavaScript pour Excel. Gr?ce ? cette m?thode, vous pouvez d?finir plusieurs propri?t?s d?un objet ? la fois en transmettant soit un autre objet du m?me type Office.js, soit un objet JavaScript avec des propri?t?s structur?es comme celles de l?objet sur lequel la m?thode est appel?e.

> [!NOTE]
> la m?thode `set()` est impl?ment?e uniquement pour les objets dans les API JavaScript pour Office propres ? un h?te, telles que l?API JavaScript pour Excel. Les API communes (partag?es) ne prennent pas en charge cette m?thode. 

### <a name="set-properties-object-options-object"></a>set (properties: object, options: object)

Les propri?t?s de l?objet sur lequel la m?thode est appel?e sont d?finies sur les valeurs sp?cifi?es par les propri?t?s correspondantes de l?objet transmis. Si le param?tre `properties` est un objet JavaScript, toute propri?t? de l?objet transmis qui correspond ? une propri?t? en lecture seule dans l?objet sur lequel la m?thode est appel?e sera ignor?e ou g?n?rera une exception, en fonction de la valeur du param?tre `options`.

#### <a name="syntax"></a>Syntaxe

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a>Param?tres

|**Param?tre**|**Type**|**Description**|
|:------------|:--------|:----------|
|`properties`|objet|Objet de m?me type Office.js que l?objet sur lequel la m?thode est appel?e ou objet JavaScript avec des noms et des types de propri?t?s refl?tant la structure de l?objet sur lequel la m?thode est appel?e.|
|`options`|objet|Facultatif. Peut ?tre transmis uniquement si le premier param?tre est un objet JavaScript. L?objet peut contenir la propri?t? suivante : `throwOnReadOnly?: boolean` (La valeur par d?faut est `true` : g?n?rer une erreur si l?objet JavaScript transmis inclut des propri?t?s en lecture seule.)|

#### <a name="returns"></a>Retourne

void    

#### <a name="example"></a>Exemple

L?exemple de code suivant d?finit plusieurs propri?t?s de mise en forme d?une plage en appelant la m?thode `set()` et en transmettant un objet JavaScript avec des noms et des types de propri?t?s refl?tant la structure des propri?t?s dans l?objet **Range**. Cet exemple part du principe que des donn?es sont pr?sentes dans la plage **B2:E2**.

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
## <a name="42ornullobject-methods"></a>M?thodes *OrNullObject

De nombreuses m?thodes d?API JavaScript pour Excel renvoient une exception lorsque la condition de l?API n?est pas remplie. Par exemple, si vous tentez d?obtenir une feuille de calcul en sp?cifiant le nom d?une feuille de calcul qui n?existe pas dans le classeur, la m?thode `getItem()` renvoie une exception `ItemNotFound`. 

Au lieu d?impl?menter une logique complexe de gestion des exceptions pour des sc?narios similaires, vous pouvez utiliser la variante de la m?thode `*OrNullObject` disponible pour les diff?rentes m?thodes de l?API JavaScript pour Excel. Une m?thode `*OrNullObject` renvoie un objet Null (pas l??l?ment JavaScript `null`) au lieu de lever une exception si l??l?ment sp?cifi? n?existe pas. Par exemple, vous pouvez appeler la m?thode `getItemOrNullObject()` sur une collection telle que **Worksheets** pour tenter de r?cup?rer un ?l?ment de la collection. La m?thode `getItemOrNullObject()` renvoie l??l?ment sp?cifi?, s?il existe. Sinon, elle renvoie un objet Null. L?objet Null renvoy? contient la propri?t? bool?enne `isNullObject` que vous pouvez ?tudier pour d?terminer l?existence de l?objet.

L?exemple de code suivant tente de r?cup?rer une feuille de calcul nomm?e ? Data ? ? l?aide de la m?thode `getItemOrNullObject()`. Si la m?thode renvoie un objet Null, une nouvelle feuille doit ?tre cr??e pour pouvoir r?aliser des actions sur la feuille.

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
 
* [Concepts de base de l?API JavaScript pour Excel](excel-add-ins-core-concepts.md)
* [Exemples de code pour les compl?ments Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Optimisation des performances de l'API JavaScript d'Excel](https://dev.office.com/reference/add-ins/excel/performance.md)
* [R?f?rence de l?API JavaScript pour Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
