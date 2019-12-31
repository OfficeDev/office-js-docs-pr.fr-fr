---
title: Appel de fonctions de feuille de calcul Excel intégrées à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 12/19/2019
localization_priority: Normal
ms.openlocfilehash: a2c98d21b36a88777e58d85c14169ffc2d67ae59
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/31/2019
ms.locfileid: "40914999"
---
# <a name="call-built-in-excel-worksheet-functions"></a>Appel de fonctions de feuille de calcul Excel intégrées

Cet article explique comment appeler les fonctions de feuille de calcul Excel intégrées telles que `VLOOKUP` et `SUM` utilisant l’API JavaScript pour Excel. Il fournit également la liste complète des fonctions de feuille de calcul Excel intégrées pouvant être appelées à l’aide de l’API JavaScript pour Excel.

> [!NOTE]
> Pour plus d’informations sur la création de *fonctions personnalisées* dans Excel à l’aide de l’API JavaScript pour Excel, reportez-vous à la rubrique [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md).

## <a name="calling-a-worksheet-function"></a>Appel d’une fonction de feuille de calcul

L’extrait de code suivant montre comment appeler une fonction de feuille de calcul où `sampleFunction()` est un espace réservé devant être remplacé par le nom de la fonction à appeler et les paramètres d’entrée nécessitant la fonction. La propriété **value** de l’objet **FunctionResult** renvoyée par une fonction de feuille de calcul contient le résultat de la fonction spécifiée. Comme le montre cet exemple, vous devez charger (`load`) la propriété **value** de l’objet **FunctionResult** avant de pouvoir la lire. Dans cet exemple, le résultat de la fonction est simplement écrit sur la console.

```js
var functionResult = context.workbook.functions.sampleFunction();
functionResult.load('value');
return context.sync()
    .then(function () {
        console.log('Result of the function: ' + functionResult.value);
    });
```

> [!TIP]
> Reportez-vous à la section [Fonctions de feuille de calcul prises en charge](#supported-worksheet-functions) de cet article pour obtenir la liste des fonctions appelées à l’aide de l’API JavaScript pour Excel.

## <a name="sample-data"></a>Exemple de données

L’image suivante montre un tableau dans une feuille de calcul Excel contenant des données de ventes pour divers types d’outils sur une période de trois mois. Chaque numéro de la table représente le nombre d’unités vendues pour un outil spécifique lors d’un mois donné. Les exemples suivant expliquent comment appliquer des fonctions de feuille de calcul intégrées à ces données.

![Capture d’écran de données de ventes dans Excel pour les catégories Hammer (Marteau), Wrench (Clé) et Saw (Scie) en novembre, décembre et janvier](../images/worksheet-functions-chaining-results.jpg)

## <a name="example-1-single-function"></a>Exemple 1 : Fonction unique

L’exemple de code suivant applique la fonction `VLOOKUP` aux exemples de données décrits précédemment pour identifier le nombre de clés vendues au mois de novembre.

```js
Excel.run(function (context) {
    var range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    var unitSoldInNov = context.workbook.functions.vlookup("Wrench", range, 2, false);
    unitSoldInNov.load('value');

    return context.sync()
        .then(function () {
            console.log(' Number of wrenches sold in November = ' + unitSoldInNov.value);
        });
}).catch(errorHandlerFunction);
```

## <a name="example-2-nested-functions"></a>Exemple 2 : Fonctions imbriquées

L’exemple de code suivant applique la fonction `VLOOKUP` pour les exemples de données décrits précédemment afin d’identifier le nombre de clés vendues au mois de novembre et le nombre de clés vendues en décembre, puis applique la fonction `SUM` pour calculer le nombre total de clés vendues au cours de ces deux mois.

Comme indiqué dans cet exemple, si un ou plusieurs appels de fonction sont imbriqués dans un autre appel de fonction, vous devez uniquement charger (`load`) le résultat final que vous souhaitez lire par la suite (dans cet exemple, `sumOfTwoLookups`). Les résultats intermédiaires (dans cet exemple, le résultat de chaque fonction `VLOOKUP`) sont calculés et utilisés pour calculer le résultat final.

```js
Excel.run(function (context) {
    var range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    var sumOfTwoLookups = context.workbook.functions.sum(
        context.workbook.functions.vlookup("Wrench", range, 2, false),
        context.workbook.functions.vlookup("Wrench", range, 3, false)
    );
    sumOfTwoLookups.load('value');

    return context.sync()
        .then(function () {
            console.log(' Number of wrenches sold in November and December = ' + sumOfTwoLookups.value);
        });
}).catch(errorHandlerFunction);
```

## <a name="supported-worksheet-functions"></a>Fonctions de feuille de calcul prises en charge

Les fonctions de feuille de calcul Excel intégrées suivantes peuvent être appelées à l’aide de l’API JavaScript pour Excel.

| Fonction | Description |
|:---------------|:-----------|
| <a href="https://support.office.com/article/ABS-function-3420200f-5628-4e8c-99da-c99d7c87713c" target="_blank">Fonction ABS</a> | Renvoie la valeur absolue d’un nombre |
| <a href="https://support.office.com/article/ACCRINT-function-fe45d089-6722-4fb3-9379-e1f911d8dc74" target="_blank">Fonction ACCRINT</a> | Renvoie l’intérêt couru non échu d’un titre dont l’intérêt est perçu périodiquement |
| <a href="https://support.office.com/article/ACCRINTM-function-f62f01f9-5754-4cc4-805b-0e70199328a7" target="_blank">Fonction ACCRINTM</a> | Renvoie l’intérêt couru non échu d’un titre dont l’intérêt est perçu à l’échéance |
| <a href="https://support.office.com/article/ACOS-function-cb73173f-d089-4582-afa1-76e5524b5d5b" target="_blank">Fonction ACOS</a> | Renvoie l’arccosinus d’un nombre |
| <a href="https://support.office.com/article/ACOSH-function-e3992cc1-103f-4e72-9f04-624b9ef5ebfe" target="_blank">Fonction ACOSH</a> | Renvoie le cosinus hyperbolique inverse d’un nombre |
| <a href="https://support.office.com/article/ACOT-function-dc7e5008-fe6b-402e-bdd6-2eea8383d905" target="_blank">Fonction ACOT</a> | Renvoie l’arccotangente d’un nombre |
| <a href="https://support.office.com/article/ACOTH-function-cc49480f-f684-4171-9fc5-73e4e852300f" target="_blank">Fonction ACOTH</a> | Renvoie l’arccotangente hyperbolique d’un nombre |
| <a href="https://support.office.com/article/AMORDEGRC-function-a14d0ca1-64a4-42eb-9b3d-b0dededf9e51" target="_blank">Fonction AMORDEGRC</a> | Renvoie l’amortissement correspondant à chaque période comptable en utilisant un coefficient d’amortissement |
| <a href="https://support.office.com/article/AMORLINC-function-7d417b45-f7f5-4dba-a0a5-3451a81079a8" target="_blank">Fonction AMORLINC</a> | Renvoie l’amortissement correspondant à chaque période comptable |
| <a href="https://support.office.com/article/AND-function-5f19b2e8-e1df-4408-897a-ce285a19e9d9" target="_blank">Fonction AND</a> | Renvoie `TRUE` si tous les arguments ont la valeur True |
| <a href="https://support.office.com/article/ARABIC-function-9a8da418-c17b-4ef9-a657-9370a30a674f" target="_blank">Fonction ARABIC</a> | Convertit un nombre romain en chiffre arabe |
| <a href="https://support.office.com/article/AREAS-function-8392ba32-7a41-43b3-96b0-3695d2ec6152" target="_blank">Fonction AREAS</a> | Renvoie le nombre de zones dans une référence |
| <a href="https://support.office.com/article/ASC-function-0b6abf1c-c663-4004-a964-ebc00b723266" target="_blank">Fonction ASC</a> | Convertit les caractères anglais pleine chasse (codés sur deux octets) ou katakana dans une chaîne de caractères en caractères à demi-chasse (codés sur un octet) |
| <a href="https://support.office.com/article/ASIN-function-81fb95e5-6d6f-48c4-bc45-58f955c6d347" target="_blank">Fonction ASIN</a> | Renvoie l’arcsinus d’un nombre |
| <a href="https://support.office.com/article/ASINH-function-4e00475a-067a-43cf-926a-765b0249717c" target="_blank">Fonction ASINH</a> | Renvoie le sinus hyperbolique inverse d’un nombre |
| <a href="https://support.office.com/article/ATAN-function-50746fa8-630a-406b-81d0-4a2aed395543" target="_blank">Fonction ATAN</a> | Renvoie l’arctangente d’un nombre |
| <a href="https://support.office.com/article/ATAN2-function-c04592ab-b9e3-4908-b428-c96b3a565033" target="_blank">Fonction ATAN2</a> | Renvoie l’arctangente des coordonnées x et y |
| <a href="https://support.office.com/article/ATANH-function-3cd65768-0de7-4f1d-b312-d01c8c930d90" target="_blank">Fonction ATANH</a> | Renvoie la tangente hyperbolique inverse d’un nombre |
| <a href="https://support.office.com/article/AVEDEV-function-58fe8d65-2a84-4dc7-8052-f3f87b5c6639" target="_blank">Fonction AVEDEV</a> | Renvoie la moyenne des écarts absolus des points de données par rapport à leur moyenne |
| <a href="https://support.office.com/article/AVERAGE-function-047bac88-d466-426c-a32b-8f33eb960cf6" target="_blank">Fonction AVERAGE</a> | Renvoie la moyenne de ses arguments |
| <a href="https://support.office.com/article/AVERAGEA-function-f5f84098-d453-4f4c-bbba-3d2c66356091" target="_blank">Fonction AVERAGEA</a> | Renvoie la moyenne de ses arguments, y compris les nombres, le texte et les valeurs logiques |
| <a href="https://support.office.com/article/AVERAGEIF-function-faec8e2e-0dec-4308-af69-f5576d8ac642" target="_blank">Fonction AVERAGEIF</a> | Renvoie la moyenne (arithmétique) de toutes les cellules d’une plage respectant un critère donné |
| <a href="https://support.office.com/article/AVERAGEIFS-function-48910c45-1fc0-4389-a028-f7c5c3001690" target="_blank">Fonction AVERAGEIFS</a> | Renvoie la moyenne (arithmétique) de toutes les cellules qui répondent à plusieurs critères |
| <a href="https://support.office.com/article/BAHTTEXT-function-5ba4d0b4-abd3-4325-8d22-7a92d59aab9c" target="_blank">Fonction BAHTTEXT</a> | Convertit un nombre en texte en utilisant le format monétaire ß (baht) |
| <a href="https://support.office.com/article/BASE-function-2ef61411-aee9-4f29-a811-1c42456c6342" target="_blank">Fonction BASE</a> | Convertit un nombre en représentation textuelle avec la base spécifiée |
| <a href="https://support.office.com/article/BESSELI-function-8d33855c-9a8d-444b-98e0-852267b1c0df" target="_blank">Fonction BESSELI</a> | Renvoie la fonction de Bessel modifiée In(x) |
| <a href="https://support.office.com/article/BESSELJ-function-839cb181-48de-408b-9d80-bd02982d94f7" target="_blank">Fonction BESSELJ</a> | Renvoie la fonction de Bessel Jn(x) |
| <a href="https://support.office.com/article/BESSELK-function-606d11bc-06d3-4d53-9ecb-2803e2b90b70" target="_blank">Fonction BESSELK</a> | Renvoie la fonction de Bessel modifiée Kn(x) |
| <a href="https://support.office.com/article/BESSELY-function-f3a356b3-da89-42c3-8974-2da54d6353a2" target="_blank">Fonction BESSELY</a> | Renvoie la fonction de Bessel Yn(x) |
| <a href="https://support.office.com/article/BETADIST-function-11188c9c-780a-42c7-ba43-9ecb5a878d31" target="_blank">Fonction BETA.DIST</a> | Renvoie la fonction de distribution cumulée suivant une loi Bêta |
| <a href="https://support.office.com/article/BETAINV-function-e84cb8aa-8df0-4cf6-9892-83a341d252eb" target="_blank">Fonction BETA.INV</a> | Renvoie l’inverse de la fonction de distribution cumulée pour une distribution bêta spécifiée |
| <a href="https://support.office.com/article/BIN2DEC-function-63905b57-b3a0-453d-99f4-647bb519cd6c" target="_blank">Fonction BIN2DEC</a> | Convertit un nombre binaire en nombre décimal |
| <a href="https://support.office.com/article/BIN2HEX-function-0375e507-f5e5-4077-9af8-28d84f9f41cc" target="_blank">Fonction BIN2HEX</a> | Convertit un nombre binaire en nombre hexadécimal |
| <a href="https://support.office.com/article/BIN2OCT-function-0a4e01ba-ac8d-4158-9b29-16c25c4c23fd" target="_blank">Fonction BIN2OCT</a> | Convertit un nombre binaire en nombre octal |
| <a href="https://support.office.com/article/BINOMDIST-function-c5ae37b6-f39c-4be2-94c2-509a1480770c" target="_blank">Fonction BINOM.DIST</a> | Renvoie la probabilité d’une variable aléatoire discrète suivant la loi binomiale |
| <a href="https://support.office.com/article/BINOMDISTRANGE-function-17331329-74c7-4053-bb4c-6653a7421595" target="_blank">Fonction BINOM.DIST.RANGE</a> | Renvoie la probabilité d’un résultat de tirage en suivant une distribution binomiale |
| <a href="https://support.office.com/article/BINOMINV-function-80a0370c-ada6-49b4-83e7-05a91ba77ac9" target="_blank">Fonction BINOM.INV</a> | Renvoie la plus petite valeur pour laquelle la distribution binomiale cumulée est inférieure ou égale à une valeur critère |
| <a href="https://support.office.com/article/BITAND-function-8a2be3d7-91c3-4b48-9517-64548008563a" target="_blank">Fonction BITAND</a> | Renvoie une opération AND au niveau du bit de deux nombres |
| <a href="https://support.office.com/article/BITLSHIFT-function-c55bb27e-cacd-4c7c-b258-d80861a03c9c" target="_blank">Fonction BITLSHIFT</a> | Renvoie un nombre décalé vers la gauche de total_décalage bits. |
| <a href="https://support.office.com/article/BITOR-function-f6ead5c8-5b98-4c9e-9053-8ad5234919b2" target="_blank">Fonction BITOR</a> | Renvoie une opération OR au niveau du bit de deux nombres |
| <a href="https://support.office.com/article/BITRSHIFT-function-274d6996-f42c-4743-abdb-4ff95351222c" target="_blank">Fonction BITRSHIFT</a> | Renvoie un nombre décalé vers la droite de total_décalage bits |
| <a href="https://support.office.com/article/BITXOR-function-c81306a1-03f9-4e89-85ac-b86c3cba10e4" target="_blank">Fonction BITXOR</a> | Renvoie une opération Exclusive Or au niveau du bit de deux nombres |
| <a href="https://support.office.com/article/CEILINGMATH-function-80f95d2f-b499-4eee-9f16-f795a8e306c8" target="_blank">Encastre. MATH, fonctions d’ECMA_CEILING</a> | Arrondit un nombre à l’entier ou au multiple supérieur le plus proche de l’argument de précision |
| <a href="https://support.office.com/article/CEILINGPRECISE-function-f366a774-527a-4c92-ba49-af0a196e66cb" target="_blank">Fonction CEILING.PRECISE</a> | Arrondit un nombre à l’entier ou au multiple le plus proche de l’argument de précision. Quel que soit le signe du nombre, le nombre est arrondi à l’unité supérieure. |
| <a href="https://support.office.com/article/CHAR-function-bbd249c8-b36e-4a91-8017-1c133f9b837a" target="_blank">Fonction CHAR</a> | Renvoie le caractère spécifié par le code numérique |
| <a href="https://support.office.com/article/CHISQDIST-function-8486b05e-5c05-4942-a9ea-f6b341518732" target="_blank">Fonction CHISQ.DIST</a> | Renvoie la fonction de densité de probabilité bêta cumulative |
| <a href="https://support.office.com/article/CHISQDISTRT-function-dc4832e8-ed2b-49ae-8d7c-b28d5804c0f2" target="_blank">Fonction CHISQ.DIST.RT</a> | Renvoie la probabilité d’une variable aléatoire continue suivant une loi unilatérale du Khi-deux |
| <a href="https://support.office.com/article/CHISQINV-function-400db556-62b3-472d-80b3-254723e7092f" target="_blank">Fonction CHISQ.INV</a> | Renvoie la fonction de densité de probabilité bêta cumulative |
| <a href="https://support.office.com/article/CHISQINVRT-function-435b5ed8-98d5-4da6-823f-293e2cbc94fe" target="_blank">Fonction CHISQ.INV.RT</a> | Renvoie l’inverse de la probabilité d’une variable aléatoire continue suivant une loi unilatérale du Khi-deux |
| <a href="https://support.office.com/article/CHOOSE-function-fc5c184f-cb62-4ec7-a46e-38653b98f5bc" target="_blank">Fonction CHOOSE</a> | Choisit une valeur dans une liste de valeurs |
| <a href="https://support.office.com/article/CLEAN-function-26f3d7c5-475f-4a9c-90e5-4b8ba987ba41" target="_blank">Fonction CLEAN</a> | Supprime tous les caractères non imprimables du texte |
| <a href="https://support.office.com/article/CODE-function-c32b692b-2ed0-4a04-bdd9-75640144b928" target="_blank">Fonction CODE</a> | Renvoie le code numérique du premier caractère d’une chaîne de texte |
| <a href="https://support.office.com/article/COLUMNS-function-4e8e7b4e-e603-43e8-b177-956088fa48ca" target="_blank">Fonction COLUMNS</a> | Renvoie le nombre de colonnes dans une référence |
| <a href="https://support.office.com/article/COMBIN-function-12a3f276-0a21-423a-8de6-06990aaf638a" target="_blank">Fonction COMBIN</a> | Renvoie le nombre de combinaisons pour un nombre d’objets donné |
| <a href="https://support.office.com/article/COMBINA-function-efb49eaa-4f4c-4cd2-8179-0ddfcf9d035d" target="_blank">Fonction COMBINA</a> | Renvoie le nombre de combinaisons avec répétitions pour un nombre d’éléments donné |
| <a href="https://support.office.com/article/COMPLEX-function-f0b8f3a9-51cc-4d6d-86fb-3a9362fa4128" target="_blank">Fonction COMPLEX</a> | Convertit des coefficients réels et imaginaires en nombre complexe |
| <a href="https://support.office.com/article/CONCATENATE-function-8f8ae884-2ca8-4f7a-b093-75d702bea31d" target="_blank">Fonction CONCATENATE</a> | Regroupe plusieurs éléments textuels en un élément textuel |
| <a href="https://support.office.com/article/CONFIDENCENORM-function-7cec58a6-85bb-488d-91c3-63828d4fbfd4" target="_blank">Fonction CONFIDENCE.NORM</a> | Renvoie l’intervalle de confiance pour la moyenne d’une population |
| <a href="https://support.office.com/article/CONFIDENCET-function-e8eca395-6c3a-4ba9-9003-79ccc61d3c53" target="_blank">Fonction CONFIDENCE.T</a> | Renvoie l’intervalle de confiance pour la moyenne d’une population, à l’aide de la probabilité d’une variable aléatoire suivant une loi T de Student |
| <a href="https://support.office.com/article/CONVERT-function-d785bef1-808e-4aac-bdcd-666c810f9af2" target="_blank">Fonction CONVERT</a> | Convertit un nombre d’un système de mesure à un autre |
| <a href="https://support.office.com/article/COS-function-0fb808a5-95d6-4553-8148-22aebdce5f05" target="_blank">Fonction COS</a> | Renvoie le cosinus d’un nombre |
| <a href="https://support.office.com/article/COSH-function-e460d426-c471-43e8-9540-a57ff3b70555" target="_blank">Fonction COSH</a> | Renvoie le cosinus hyperbolique d’un nombre |
| <a href="https://support.office.com/article/COT-function-c446f34d-6fe4-40dc-84f8-cf59e5f5e31a" target="_blank">Fonction COT</a> | Renvoie la cotangente d’un angle |
| <a href="https://support.office.com/article/COTH-function-2e0b4cb6-0ba0-403e-aed4-deaa71b49df5" target="_blank">Fonction COTH</a> | Renvoie la cotangente hyperbolique d’un nombre |
| <a href="https://support.office.com/article/COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c" target="_blank">Fonction COUNT</a> | Compte le nombre de chiffres compris dans la liste d’arguments |
| <a href="https://support.office.com/article/COUNTA-function-7dc98875-d5c1-46f1-9a82-53f3219e2509" target="_blank">Fonction COUNTA</a> | Compte le nombre de valeurs comprises dans la liste d’arguments |
| <a href="https://support.office.com/article/COUNTBLANK-function-6a92d772-675c-4bee-b346-24af6bd3ac22" target="_blank">Fonction COUNTBLANK</a> | Compte le nombre de cellules vides dans une plage |
| <a href="https://support.office.com/article/COUNTIF-function-e0de10c6-f885-4e71-abb4-1f464816df34" target="_blank">Fonction COUNTIF</a> | Compte le nombre de cellules à l’intérieur d’une plage qui répondent aux critères donnés |
| <a href="https://support.office.com/article/COUNTIFS-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842" target="_blank">Fonction COUNTIFS</a> | Compte le nombre de cellules à l’intérieur d’une plage qui répondent à plusieurs critères |
| <a href="https://support.office.com/article/COUPDAYBS-function-eb9a8dfb-2fb2-4c61-8e5d-690b320cf872" target="_blank">Fonction COUPDAYBS</a> | Renvoie le nombre de jours entre le début de la période du coupon et la date d’escompte |
| <a href="https://support.office.com/article/COUPDAYS-function-cc64380b-315b-4e7b-950c-b30b0a76f671" target="_blank">Fonction COUPDAYS</a> | Renvoie le nombre de jours dans la période du coupon contenant la date d’escompte |
| <a href="https://support.office.com/article/COUPDAYSNC-function-5ab3f0b2-029f-4a8b-bb65-47d525eea547" target="_blank">Fonction COUPDAYSNC</a> | Renvoie le nombre de jours séparant la date d’escompte de la date du prochain coupon |
| <a href="https://support.office.com/article/COUPNCD-function-fd962fef-506b-4d9d-8590-16df5393691f" target="_blank">Fonction COUPNCD</a> | Renvoie la date du prochain coupon suivant la date d’escompte |
| <a href="https://support.office.com/article/COUPNUM-function-a90af57b-de53-4969-9c99-dd6139db2522" target="_blank">Fonction COUPNUM</a> | Renvoie le nombre de coupons à régler entre la date d’escompte et la date d’échéance |
| <a href="https://support.office.com/article/COUPPCD-function-2eb50473-6ee9-4052-a206-77a9a385d5b3" target="_blank">Fonction COUPPCD</a> | Renvoie la date du coupon antérieur précédant la date d’escompte |
| <a href="https://support.office.com/article/CSC-function-07379361-219a-4398-8675-07ddc4f135c1" target="_blank">Fonction CSC</a> | Renvoie la cosécante d’un angle |
| <a href="https://support.office.com/article/CSCH-function-f58f2c22-eb75-4dd6-84f4-a503527f8eeb" target="_blank">Fonction CSCH</a> | Renvoie la cosécante hyperbolique d’un angle |
| <a href="https://support.office.com/article/CUMIPMT-function-61067bb0-9016-427d-b95b-1a752af0e606" target="_blank">Fonction CUMIPMT</a> | Renvoie les intérêts cumulés réglés entre deux périodes |
| <a href="https://support.office.com/article/CUMPRINC-function-94a4516d-bd65-41a1-bc16-053a6af4c04d" target="_blank">Fonction CUMPRINC</a> | Renvoie le montant cumulé du remboursement du capital réglé entre deux périodes |
| <a href="https://support.office.com/article/DATE-function-e36c0c8c-4104-49da-ab83-82328b832349" target="_blank">Fonction DATE</a> | Renvoie le numéro de série d’une date précise |
| <a href="https://support.office.com/article/DATEVALUE-function-df8b07d4-7761-4a93-bc33-b7471bbff252" target="_blank">Fonction DATEVALUE</a> | Convertit une date au format texte en numéro de série |
| <a href="https://support.office.com/article/DAVERAGE-function-a6a2d5ac-4b4b-48cd-a1d8-7b37834e5aee" target="_blank">Fonction DAVERAGE</a> | Renvoie la moyenne des entrées d’une base de données sélectionnée |
| <a href="https://support.office.com/article/DAY-function-8a7d1cbb-6c7d-4ba1-8aea-25c134d03101" target="_blank">Fonction DAY</a> | Convertit un numéro de série en jour du mois |
| <a href="https://support.office.com/article/DAYS-function-57740535-d549-4395-8728-0f07bff0b9df" target="_blank">Fonction DAYS</a> | Renvoie le nombre de jours entre deux dates |
| <a href="https://support.office.com/article/DAYS360-function-b9a509fd-49ef-407e-94df-0cbda5718c2a" target="_blank">Fonction DAYS360</a> | Calcule le nombre de jours entre deux dates sur la base d’une année de 360 jours |
| <a href="https://support.office.com/article/DB-function-354e7d28-5f93-4ff1-8a52-eb4ee549d9d7" target="_blank">Fonction DB</a> | Renvoie l’amortissement d’un bien durant une période spécifiée en utilisant la méthode de l’amortissement dégressif à taux fixe |
| <a href="https://support.office.com/article/DBCS-function-a4025e73-63d2-4958-9423-21a24794c9e5" target="_blank">Fonction DBCS</a> | Convertit les caractères anglais à demi-chasse (codés sur un octet) ou katakana dans une chaîne de caractères en caractères pleine chasse (codés sur deux octets) |
| <a href="https://support.office.com/article/DCOUNT-function-c1fc7b93-fb0d-4d8d-97db-8d5f076eaeb1" target="_blank">Fonction DCOUNT</a> | Compte les cellules qui contiennent des nombres dans une base de données |
| <a href="https://support.office.com/article/DCOUNTA-function-00232a6d-5a66-4a01-a25b-c1653fda1244" target="_blank">Fonction DCOUNTA</a> | Compte les cellules non vides d’une base de données |
| <a href="https://support.office.com/article/DDB-function-519a7a37-8772-4c96-85c0-ed2c209717a5" target="_blank">Fonction DDB</a> | Renvoie l’amortissement d’un bien durant une période spécifiée suivant la méthode de l’amortissement dégressif à taux double ou selon un coefficient à spécifier |
| <a href="https://support.office.com/article/DEC2BIN-function-0f63dd0e-5d1a-42d8-b511-5bf5c6d43838" target="_blank">Fonction DEC2BIN</a> | Convertit un nombre décimal en nombre binaire |
| <a href="https://support.office.com/article/DEC2HEX-function-6344ee8b-b6b5-4c6a-a672-f64666704619" target="_blank">Fonction DEC2HEX</a> | Convertit un nombre décimal en nombre hexadécimal |
| <a href="https://support.office.com/article/DEC2OCT-function-c9d835ca-20b7-40c4-8a9e-d3be351ce00f" target="_blank">Fonction DEC2OCT</a> | Convertit un nombre décimal en nombre octal |
| <a href="https://support.office.com/article/DECIMAL-function-ee554665-6176-46ef-82de-0a283658da2e" target="_blank">Fonction DECIMAL</a> | Convertit une représentation textuelle d’un nombre dans une base donnée en nombre décimal |
| <a href="https://support.office.com/article/DEGREES-function-4d6ec4db-e694-4b94-ace0-1cc3f61f9ba1" target="_blank">Fonction DEGREES</a> | Convertit des radians en degrés |
| <a href="https://support.office.com/article/DELTA-function-2f763672-c959-4e07-ac33-fe03220ba432" target="_blank">Fonction DELTA</a> | Vérifie si deux valeurs sont égales |
| <a href="https://support.office.com/article/DEVSQ-function-8b739616-8376-4df5-8bd0-cfe0a6caf444" target="_blank">Fonction DEVSQ</a> | Renvoie la somme des carrés des écarts |
| <a href="https://support.office.com/article/DGET-function-455568bf-4eef-45f7-90f0-ec250d00892e" target="_blank">Fonction DGET</a> | Extrait d’une base de données un seul enregistrement correspondant aux critères spécifiés |
| <a href="https://support.office.com/article/DISC-function-71fce9f3-3f05-4acf-a5a3-eac6ef4daa53" target="_blank">Fonction DISC</a> | Renvoie le taux d’escompte d’un titre |
| <a href="https://support.office.com/article/DMAX-function-f4e8209d-8958-4c3d-a1ee-6351665d41c2" target="_blank">Fonction DMAX</a> | Renvoie la valeur maximale des entrées de base de données sélectionnées |
| <a href="https://support.office.com/article/DMIN-function-4ae6f1d9-1f26-40f1-a783-6dc3680192a3" target="_blank">Fonction DMIN</a> | Renvoie la valeur minimale des entrées de base de données sélectionnées |
| <a href="https://support.office.com/article/DOLLAR-function-a6cd05d9-9740-4ad3-a469-8109d18ff611" target="_blank">DOLLAR, fonctions USDollar,</a> | Convertit un nombre en texte en utilisant le format monétaire $ (dollar) |
| <a href="https://support.office.com/article/DOLLARDE-function-db85aab0-1677-428a-9dfd-a38476693427" target="_blank">Fonction DOLLARDE</a> | Convertit un prix en dollars, exprimé sous forme de fraction, en un prix en dollars exprimé sous forme de nombre décimal |
| <a href="https://support.office.com/article/DOLLARFR-function-0835d163-3023-4a33-9824-3042c5d4f495" target="_blank">Fonction DOLLARFR</a> | Convertit un prix en dollars, exprimé sous forme de nombre décimal, en un prix en dollars exprimé sous forme de fraction |
| <a href="https://support.office.com/article/DPRODUCT-function-4f96b13e-d49c-47a7-b769-22f6d017cb31" target="_blank">Fonction DPRODUCT</a> | Multiplie les valeurs d’un champ particulier dans des enregistrements correspondant aux critères d’une base de données |
| <a href="https://support.office.com/article/DSTDEV-function-026b8c73-616d-4b5e-b072-241871c4ab96" target="_blank">Fonction DSTDEV</a> | Calcule l’écart type en fonction d’un échantillon d’entrées de base de données sélectionnées |
| <a href="https://support.office.com/article/DSTDEVP-function-04b78995-da03-4813-bbd9-d74fd0f5d94b" target="_blank">Fonction DSTDEVP</a> | Calcule l’écart type en fonction de l’ensemble des entrées de base de données sélectionnées |
| <a href="https://support.office.com/article/DSUM-function-53181285-0c4b-4f5a-aaa3-529a322be41b" target="_blank">Fonction DSUM</a> | Ajoute les nombres dans la colonne Champ des enregistrements de la base de données correspondant aux critères |
| <a href="https://support.office.com/article/DURATION-function-b254ea57-eadc-4602-a86a-c8e369334038" target="_blank">Fonction DURATION</a> | Renvoie la durée annuelle d’un titre dont les intérêts sont perçus périodiquement |
| <a href="https://support.office.com/article/DVAR-function-d6747ca9-99c7-48bb-996e-9d7af00f3ed1" target="_blank">Fonction DVAR</a> | Estime la variance en fonction d’un échantillon d’entrées de base de données sélectionnées |
| <a href="https://support.office.com/article/DVARP-function-eb0ba387-9cb7-45c8-81e9-0394912502fc" target="_blank">Fonction DVARP</a> | Calcule la variance en fonction de l’ensemble des entrées de base de données sélectionnées |
| <a href="https://support.office.com/article/EDATE-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5" target="_blank">Fonction EDATE</a> | Renvoie le numéro de série de la date qui représente le nombre indiqué de mois précédant ou suivant la date de début |
| <a href="https://support.office.com/article/EFFECT-function-910d4e4c-79e2-4009-95e6-507e04f11bc4" target="_blank">Fonction EFFECT</a> | Renvoie le taux d’intérêt annuel effectif |
| <a href="https://support.office.com/article/EOMONTH-function-7314ffa1-2bc9-4005-9d66-f49db127d628" target="_blank">Fonction EOMONTH</a> | Renvoie le numéro de série du dernier jour du mois précédant ou suivant un nombre de mois spécifié |
| <a href="https://support.office.com/article/ERF-function-c53c7e7b-5482-4b6c-883e-56df3c9af349" target="_blank">Fonction ERF</a> | Renvoie la valeur de la fonction d’erreur |
| <a href="https://support.office.com/article/ERFPRECISE-function-9a349593-705c-4278-9a98-e4122831a8e0" target="_blank">Fonction ERF.PRECISE</a> | Renvoie la valeur de la fonction d’erreur |
| <a href="https://support.office.com/article/ERFC-function-736e0318-70ba-4e8b-8d08-461fe68b71b3" target="_blank">Fonction ERFC</a> | Renvoie la valeur de la fonction d’erreur complémentaire |
| <a href="https://support.office.com/article/ERFCPRECISE-function-e90e6bab-f45e-45df-b2ac-cd2eb4d4a273" target="_blank">Fonction ERFC.PRECISE</a> | Renvoie la valeur de la fonction d’erreur complémentaire comprise entre x et l’infini |
| <a href="https://support.office.com/article/ERRORTYPE-function-10958677-7c8d-44f7-ae77-b9a9ee6eefaa" target="_blank">Fonction ERROR.TYPE</a> | Renvoie un nombre correspondant à un type d’erreur |
| <a href="https://support.office.com/article/EVEN-function-197b5f06-c795-4c1e-8696-3c3b8a646cf9" target="_blank">Fonction EVEN</a> | Arrondit un nombre au nombre entier pair supérieur |
| <a href="https://support.office.com/article/EXACT-function-d3087698-fc15-4a15-9631-12575cf29926" target="_blank">Fonction EXACT</a> | Vérifie si deux valeurs textuelles sont identiques |
| <a href="https://support.office.com/article/EXP-function-c578f034-2c45-4c37-bc8c-329660a63abe" target="_blank">Fonction EXP</a> | Renvoie le nombre e élevé à la puissance d’un nombre donné |
| <a href="https://support.office.com/article/EXPONDIST-function-4c12ae24-e563-4155-bf3e-8b78b6ae140e" target="_blank">Fonction EXPON.DIST</a> | Renvoie la distribution exponentielle |
| <a href="https://support.office.com/article/FDIST-function-a887efdc-7c8e-46cb-a74a-f884cd29b25d" target="_blank">Fonction F.DIST</a> | Renvoie la distribution de probabilité F |
| <a href="https://support.office.com/article/FDISTRT-function-d74cbb00-6017-4ac9-b7d7-6049badc0520" target="_blank">Fonction F.DIST.RT</a> | Renvoie la distribution de probabilité F |
| <a href="https://support.office.com/article/FINV-function-0dda0cf9-4ea0-42fd-8c3c-417a1ff30dbe" target="_blank">Fonction F.INV</a> | Renvoie l’inverse de la distribution de probabilité F |
| <a href="https://support.office.com/article/FINVRT-function-d371aa8f-b0b1-40ef-9cc2-496f0693ac00" target="_blank">Fonction F.INVERT.RT</a> | Renvoie l’inverse de la distribution de probabilité F |
| <a href="https://support.office.com/article/FACT-function-ca8588c2-15f2-41c0-8e8c-c11bd471a4f3" target="_blank">Fonction FACT</a> | Renvoie la factorielle d’un nombre |
| <a href="https://support.office.com/article/FACTDOUBLE-function-e67697ac-d214-48eb-b7b7-cce2589ecac8" target="_blank">Fonction FACTDOUBLE</a> | Renvoie la factorielle double d’un nombre |
| <a href="https://support.office.com/article/FALSE-function-2d58dfa5-9c03-4259-bf8f-f0ae14346904" target="_blank">Fonction FALSE</a> | Renvoie la valeur logique `FALSE` |
| <a href="https://support.office.com/article/FIND-FINDB-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628" target="_blank">Fonctions FIND, FINDB</a> | Cherche une valeur textuelle dans une autre (en respectant la casse) |
| <a href="https://support.office.com/article/FISHER-function-d656523c-5076-4f95-b87b-7741bf236c69" target="_blank">Fonction FISHER</a> | Renvoie la transformation de Fisher |
| <a href="https://support.office.com/article/FISHERINV-function-62504b39-415a-4284-a285-19c8e82f86bb" target="_blank">Fonction FISHERINV</a> | Renvoie l’inverse de la transformation de Fisher |
| <a href="https://support.office.com/article/FIXED-function-ffd5723c-324c-45e9-8b96-e41be2a8274a" target="_blank">Fonction FIXED</a> | Convertit un nombre en texte avec un nombre de décimales fixe |
| <a href="https://support.office.com/article/FLOORMATH-function-c302b599-fbdb-4177-ba19-2c2b1249a2f5" target="_blank">Fonction FLOOR.MATH</a> | Arrondit un nombre à l’entier ou au multiple inférieur le plus proche de l’argument de précision |
| <a href="https://support.office.com/article/FLOORPRECISE-function-f769b468-1452-4617-8dc3-02f842a0702e" target="_blank">Fonction FLOOR.PRECISE</a> | Arrondit un nombre à l’entier ou au multiple inférieur le plus proche de l’argument de précision. Quel que soit le signe du nombre, le nombre est arrondi à l’unité inférieure. |
| <a href="https://support.office.com/article/FV-function-2eef9f44-a084-4c61-bdd8-4fe4bb1b71b3" target="_blank">Fonction FV</a> | Renvoie la valeur future d’un investissement |
| <a href="https://support.office.com/article/FVSCHEDULE-function-bec29522-bd87-4082-bab9-a241f3fb251d" target="_blank">Fonction FVSCHEDULE</a> | Renvoie la valeur future d’un investissement en appliquant une série de taux d’intérêt composites |
| <a href="https://support.office.com/article/GAMMA-function-ce1702b1-cf55-471d-8307-f83be0fc5297" target="_blank">Fonction GAMMA</a> | Renvoie la valeur de la fonction Gamma |
| <a href="https://support.office.com/article/GAMMADIST-function-9b6f1538-d11c-4d5f-8966-21f6a2201def" target="_blank">Fonction GAMMA.DIST</a> | Renvoie la distribution suivant une loi Gamma |
| <a href="https://support.office.com/article/GAMMAINV-function-74991443-c2b0-4be5-aaab-1aa4d71fbb18" target="_blank">Fonction GAMMA.INV</a> | Renvoie l’inverse de la distribution cumulée suivant une loi Gamma |
| <a href="https://support.office.com/article/GAMMALN-function-b838c48b-c65f-484f-9e1d-141c55470eb9" target="_blank">Fonction GAMMALN</a> | Renvoie le logarithme népérien de la fonction gamma, Γ(x) |
| <a href="https://support.office.com/article/GAMMALNPRECISE-function-5cdfe601-4e1e-4189-9d74-241ef1caa599" target="_blank">Fonction GAMMALN.PRECISE</a> | Renvoie le logarithme népérien de la fonction gamma, Γ(x) |
| <a href="https://support.office.com/article/GAUSS-function-069f1b4e-7dee-4d6a-a71f-4b69044a6b33" target="_blank">Fonction GAUSS</a> | Renvoie 0,5 de moins que la distribution cumulée suivant une loi normale centrée réduite |
| <a href="https://support.office.com/article/GCD-function-d5107a51-69e3-461f-8e4c-ddfc21b5073a" target="_blank">Fonction GCD</a> | Renvoie le plus grand diviseur commun |
| <a href="https://support.office.com/article/GEOMEAN-function-db1ac48d-25a5-40a0-ab83-0b38980e40d5" target="_blank">Fonction GEOMEAN</a> | Renvoie la moyenne géométrique |
| <a href="https://support.office.com/article/GESTEP-function-f37e7d2a-41da-4129-be95-640883fca9df" target="_blank">Fonction GESTEP</a> | Vérifie si un nombre est supérieur à une valeur seuil |
| <a href="https://support.office.com/article/HARMEAN-function-5efd9184-fab5-42f9-b1d3-57883a1d3bc6" target="_blank">Fonction HARMEAN</a> | Renvoie la moyenne harmonique |
| <a href="https://support.office.com/article/HEX2BIN-function-a13aafaa-5737-4920-8424-643e581828c1" target="_blank">Fonction HEX2BIN</a> | Convertit un nombre hexadécimal en nombre binaire |
| <a href="https://support.office.com/article/HEX2DEC-function-8c8c3155-9f37-45a5-a3ee-ee5379ef106e" target="_blank">Fonction HEX2DEC</a> | Convertit un nombre hexadécimal en nombre décimal |
| <a href="https://support.office.com/article/HEX2OCT-function-54d52808-5d19-4bd0-8a63-1096a5d11912" target="_blank">Fonction HEX2OCT</a> | Convertit un nombre hexadécimal en nombre octal |
| <a href="https://support.office.com/article/HLOOKUP-function-a3034eec-b719-4ba3-bb65-e1ad662ed95f" target="_blank">Fonction HLOOKUP</a> | Cherche dans la première ligne d’un tableau et renvoie la valeur de la cellule indiquée |
| <a href="https://support.office.com/article/HOUR-function-a3afa879-86cb-4339-b1b5-2dd2d7310ac7" target="_blank">Fonction HOUR</a> | Convertit un numéro de série en heure |
| <a href="https://support.office.com/article/HYPERLINK-function-333c7ce6-c5ae-4164-9c47-7de9b76f577f" target="_blank">Fonction HYPERLINK</a> | Crée un raccourci ou un renvoi qui ouvre un document stocké sur un serveur réseau, un intranet ou Internet |
| <a href="https://support.office.com/article/HYPGEOMDIST-function-6dbd547f-1d12-4b1f-8ae5-b0d9e3d22fbf" target="_blank">Fonction HYPGEOM.DIST</a> | Renvoie la distribution suivant une loi hypergéométrique |
| <a href="https://support.office.com/article/IF-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2" target="_blank">Fonction IF</a> | Indique un test logique à effectuer |
| <a href="https://support.office.com/article/IMABS-function-b31e73c6-d90c-4062-90bc-8eb351d765a1" target="_blank">Fonction IMABS</a> | Renvoie la valeur absolue (module) d’un nombre complexe |
| <a href="https://support.office.com/article/IMAGINARY-function-dd5952fd-473d-44d9-95a1-9a17b23e428a" target="_blank">Fonction IMAGINARY</a> | Renvoie le coefficient imaginaire d’un nombre complexe |
| <a href="https://support.office.com/article/IMARGUMENT-function-eed37ec1-23b3-4f59-b9f3-d340358a034a" target="_blank">Fonction IMARGUMENT</a> | Renvoie l’argument thêta, un angle exprimé en radians |
| <a href="https://support.office.com/article/IMCONJUGATE-function-2e2fc1ea-f32b-4f9b-9de6-233853bafd42" target="_blank">Fonction IMCONJUGATE</a> | Renvoie le conjugué complexe d’un nombre complexe |
| <a href="https://support.office.com/article/IMCOS-function-dad75277-f592-4a6b-ad6c-be93a808a53c" target="_blank">Fonction IMCOS</a> | Renvoie le cosinus d’un nombre complexe |
| <a href="https://support.office.com/article/IMCOSH-function-053e4ddb-4122-458b-be9a-457c405e90ff" target="_blank">Fonction IMCOSH</a> | Renvoie le cosinus hyperbolique d’un nombre complexe |
| <a href="https://support.office.com/article/IMCOT-function-dc6a3607-d26a-4d06-8b41-8931da36442c" target="_blank">Fonction IMCOT</a> | Renvoie la cotangente d’un nombre complexe |
| <a href="https://support.office.com/article/IMCSC-function-9e158d8f-2ddf-46cd-9b1d-98e29904a323" target="_blank">Fonction IMCSC</a> | Renvoie la cosécante d’un nombre complexe |
| <a href="https://support.office.com/article/IMCSCH-function-c0ae4f54-5f09-4fef-8da0-dc33ea2c5ca9" target="_blank">Fonction IMCSCH</a> | Renvoie la cosécante hyperbolique d’un nombre complexe |
| <a href="https://support.office.com/article/IMDIV-function-a505aff7-af8a-4451-8142-77ec3d74d83f" target="_blank">Fonction IMDIV</a> | Renvoie le quotient de deux nombres complexes |
| <a href="https://support.office.com/article/IMEXP-function-c6f8da1f-e024-4c0c-b802-a60e7147a95f" target="_blank">Fonction IMEXP</a> | Renvoie la fonction exponentielle d’un nombre complexe |
| <a href="https://support.office.com/article/IMLN-function-32b98bcf-8b81-437c-a636-6fb3aad509d8" target="_blank">Fonction IMLN</a> | Renvoie le logarithme népérien d’un nombre complexe |
| <a href="https://support.office.com/article/IMLOG10-function-58200fca-e2a2-4271-8a98-ccd4360213a5" target="_blank">Fonction IMLOG10</a> | Calcule le logarithme d’un nombre complexe en base 10 |
| <a href="https://support.office.com/article/IMLOG2-function-152e13b4-bc79-486c-a243-e6a676878c51" target="_blank">Fonction IMLOG2</a> | Calcule le logarithme d’un nombre complexe en base 2 |
| <a href="https://support.office.com/article/IMPOWER-function-210fd2f5-f8ff-4c6a-9d60-30e34fbdef39" target="_blank">Fonction IMPOWER</a> | Renvoie un nombre complexe élevé à une puissance entière |
| <a href="https://support.office.com/article/IMPRODUCT-function-2fb8651a-a4f2-444f-975e-8ba7aab3a5ba" target="_blank">Fonction IMPRODUCT</a> | Renvoie le produit de 2 à 255 nombres complexes |
| <a href="https://support.office.com/article/IMREAL-function-d12bc4c0-25d0-4bb3-a25f-ece1938bf366" target="_blank">Fonction IMREAL</a> | Renvoie le coefficient réel d’un nombre complexe |
| <a href="https://support.office.com/article/IMSEC-function-6df11132-4411-4df4-a3dc-1f17372459e0" target="_blank">Fonction IMSEC</a> | Renvoie la sécante d’un nombre complexe |
| <a href="https://support.office.com/article/IMSECH-function-f250304f-788b-4505-954e-eb01fa50903b" target="_blank">Fonction IMSECH</a> | Renvoie la sécante hyperbolique d’un nombre complexe |
| <a href="https://support.office.com/article/IMSIN-function-1ab02a39-a721-48de-82ef-f52bf37859f6" target="_blank">Fonction IMSIN</a> | Renvoie le sinus d’un nombre complexe |
| <a href="https://support.office.com/article/IMSINH-function-dfb9ec9e-8783-4985-8c42-b028e9e8da3d" target="_blank">Fonction IMSINH</a> | Renvoie le sinus hyperbolique d’un nombre complexe |
| <a href="https://support.office.com/article/IMSQRT-function-e1753f80-ba11-4664-a10e-e17368396b70" target="_blank">Fonction IMSQRT</a> | Renvoie la racine carrée d’un nombre complexe |
| <a href="https://support.office.com/article/IMSUB-function-2e404b4d-4935-4e85-9f52-cb08b9a45054" target="_blank">Fonction IMSUB</a> | Renvoie la différence entre deux nombres complexes |
| <a href="https://support.office.com/article/IMSUM-function-81542999-5f1c-4da6-9ffe-f1d7aaa9457f" target="_blank">Fonction IMSUM</a> | Renvoie la somme de plusieurs nombres complexes |
| <a href="https://support.office.com/article/IMTAN-function-8478f45d-610a-43cf-8544-9fc0b553a132" target="_blank">Fonction IMTAN</a> | Renvoie la tangente d’un nombre complexe |
| <a href="https://support.office.com/article/INT-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef" target="_blank">Fonction INT</a> | Arrondit un nombre à l’entier inférieur le plus proche |
| <a href="https://support.office.com/article/INTRATE-function-5cb34dde-a221-4cb6-b3eb-0b9e55e1316f" target="_blank">Fonction INTRATE</a> | Renvoie le taux d’intérêt pour un titre totalement investi |
| <a href="https://support.office.com/article/IPMT-function-5cce0ad6-8402-4a41-8d29-61a0b054cb6f" target="_blank">Fonction IPMT</a> | Renvoie le montant des intérêts d’un investissement pour une période donnée |
| <a href="https://support.office.com/article/IRR-function-64925eaa-9988-495b-b290-3ad0c163c1bc" target="_blank">Fonction IRR</a> | Renvoie le taux de rendement interne pour une série de mouvements de trésorerie |
| <a href="https://support.office.com/article/ISERR-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fonction ISERR</a> | Renvoie `TRUE` si la valeur est une valeur d’erreur, sauf #N/A |
| <a href="https://support.office.com/article/ISERROR-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fonction ISERROR</a> | Renvoie `TRUE` si la valeur est une valeur d’erreur |
| <a href="https://support.office.com/article/ISEVEN-function-aa15929a-d77b-4fbb-92f4-2f479af55356" target="_blank">Fonction ISEVEN</a> | Renvoie `TRUE` si le nombre est pair |
| <a href="https://support.office.com/article/ISFORMULA-function-e4d1355f-7121-4ef2-801e-3839bfd6b1e5" target="_blank">Fonction ISFORMULA</a> | Renvoie `TRUE` s’il existe une référence à une cellule qui contient une formule |
| <a href="https://support.office.com/article/ISLOGICAL-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fonction ISLOGICAL</a> | Renvoie `TRUE` si la valeur est une valeur logique |
| <a href="https://support.office.com/article/ISNA-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fonction ISNA</a> | Renvoie `TRUE` si la valeur est la valeur d’erreur #N/A |
| <a href="https://support.office.com/article/ISNONTEXT-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fonction ISNONTEXT</a> | Renvoie `TRUE` si la valeur n’est pas textuelle |
| <a href="https://support.office.com/article/ISNUMBER-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fonction ISNUMBER</a> | Renvoie `TRUE` si la valeur est un nombre |
| <a href="https://support.office.com/article/ISOCEILING-function-e587bb73-6cc2-4113-b664-ff5b09859a83" target="_blank">Fonction ISO.CEILING</a> | Renvoie un nombre arrondi à l’entier ou au multiple supérieur le plus proche de l’argument de précision |
| <a href="https://support.office.com/article/ISODD-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fonction ISODD</a> | Renvoie `TRUE` si le nombre est impair |
| <a href="https://support.office.com/article/ISOWEEKNUM-function-1c2d0afe-d25b-4ab1-8894-8d0520e90e0e" target="_blank">Fonction ISOWEEKNUM</a> | Renvoie le numéro de la semaine ISO de l’année pour une date donnée |
| <a href="https://support.office.com/article/ISPMT-function-fa58adb6-9d39-4ce0-8f43-75399cea56cc" target="_blank">Fonction ISPMT</a> | Calcule le montant des intérêts payés au cours d’une période spécifique d’un investissement |
| <a href="https://support.office.com/article/ISREF-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fonction ISREF</a> | Renvoie `TRUE` si la valeur est une référence |
| <a href="https://support.office.com/article/ISTEXT-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fonction ISTEXT</a> | Renvoie `TRUE` si la valeur est textuelle |
| <a href="https://support.office.com/article/KURT-function-bc3a265c-5da4-4dcb-b7fd-c237789095ab" target="_blank">Fonction KURT</a> | Renvoie le kurtosis d’un jeu de données |
| <a href="https://support.office.com/article/LARGE-function-3af0af19-1190-42bb-bb8b-01672ec00a64" target="_blank">Fonction LARGE</a> | Renvoie la k-ième plus grande valeur d’un jeu de données |
| <a href="https://support.office.com/article/LCM-function-7152b67a-8bb5-4075-ae5c-06ede5563c94" target="_blank">Fonction LCM</a> | Renvoie le plus petit dénominateur commun |
| <a href="https://support.office.com/article/LEFT-LEFTB-functions-9203d2d2-7960-479b-84c6-1ea52b99640c" target="_blank">Fonctions LEFT, LEFTB</a> | Renvoie les caractères les plus à gauche d’une valeur textuelle |
| <a href="https://support.office.com/article/LEN-LENB-functions-29236f94-cedc-429d-affd-b5e33d2c67cb" target="_blank">Fonctions LEN, LENB</a> | Renvoie le nombre de caractères dans une chaîne de texte |
| <a href="https://support.office.com/article/LN-function-81fe1ed7-dac9-4acd-ba1d-07a142c6118f" target="_blank">Fonction LN</a> | Renvoie le logarithme népérien d’un nombre |
| <a href="https://support.office.com/article/LOG-function-4e82f196-1ca9-4747-8fb0-6c4a3abb3280" target="_blank">Fonction LOG</a> | Renvoie le logarithme d’un nombre selon la base spécifiée |
| <a href="https://support.office.com/article/LOG10-function-c75b881b-49dd-44fb-b6f4-37e3486a0211" target="_blank">Fonction LOG10</a> | Renvoie le logarithme d’un nombre en base 10 |
| <a href="https://support.office.com/article/LOGNORMDIST-function-eb60d00b-48a9-4217-be2b-6074aee6b070" target="_blank">Fonction LOGNORM.DIST</a> | Renvoie la distribution suivant une loi lognormale cumulée |
| <a href="https://support.office.com/article/LOGNORMINV-function-fe79751a-f1f2-4af8-a0a1-e151b2d4f600" target="_blank">Fonction LOGNORM.INV</a> | Renvoie l’inverse de la distribution cumulée suivant une loi lognormale |
| <a href="https://support.office.com/article/LOOKUP-function-446d94af-663b-451d-8251-369d5e3864cb" target="_blank">Fonction LOOKUP</a> | Cherche des valeurs dans un vecteur ou un tableau |
| <a href="https://support.office.com/article/LOWER-function-3f21df02-a80c-44b2-afaf-81358f9fdeb4" target="_blank">Fonction LOWER</a> | Convertit le texte en minuscules |
| <a href="https://support.office.com/article/MATCH-function-e8dffd45-c762-47d6-bf89-533f4a37673a" target="_blank">Fonction MATCH</a> | Cherche des valeurs dans une référence ou un tableau |
| <a href="https://support.office.com/article/MAX-function-e0012414-9ac8-4b34-9a47-73e662c08098" target="_blank">Fonction MAX</a> | Renvoie la valeur maximale contenue dans une liste d’arguments |
| <a href="https://support.office.com/article/MAXA-function-814bda1e-3840-4bff-9365-2f59ac2ee62d" target="_blank">Fonction MAXA</a> | Renvoie la valeur maximale contenue dans une liste d’arguments, y compris les nombres, le texte et les valeurs logiques |
| <a href="https://support.office.com/article/MDURATION-function-b3786a69-4f20-469a-94ad-33e5b90a763c" target="_blank">Fonction MDURATION</a> | Renvoie la durée modifiée de Macauley pour un titre avec une valeur estimée à 100 dollars |
| <a href="https://support.office.com/article/MEDIAN-function-d0916313-4753-414c-8537-ce85bdd967d2" target="_blank">Fonction MEDIAN</a> | Renvoie la valeur médiane des nombres donnés |
| <a href="https://support.office.com/article/MID-MIDB-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028" target="_blank">Fonction MID, MIDB</a> | Renvoie un nombre déterminé de caractères d’une chaîne de texte en commençant à la position indiquée |
| <a href="https://support.office.com/article/MIN-function-61635d12-920f-4ce2-a70f-96f202dcc152" target="_blank">Fonction MIN</a> | Renvoie la valeur minimale contenue dans une liste d’arguments |
| <a href="https://support.office.com/article/MINA-function-245a6f46-7ca5-4dc7-ab49-805341bc31d3" target="_blank">Fonction MINA</a> | Renvoie la plus petite valeur contenue dans une liste d’arguments, y compris les nombres, le texte et les valeurs logiques |
| <a href="https://support.office.com/article/MINUTE-function-af728df0-05c4-4b07-9eed-a84801a60589" target="_blank">Fonction MINUTE</a> | Convertit un numéro de série en minute |
| <a href="https://support.office.com/article/MIRR-function-b020f038-7492-4fb4-93c1-35c345b53524" target="_blank">Fonction MIRR</a> | Renvoie le taux de rendement interne lorsque des mouvements de trésorerie positifs et négatifs sont financés à des taux différents |
| <a href="https://support.office.com/article/MOD-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3" target="_blank">Fonction MOD</a> | Renvoie le reste d’une division |
| <a href="https://support.office.com/article/MONTH-function-579a2881-199b-48b2-ab90-ddba0eba86e8" target="_blank">Fonction MONTH</a> | Convertit un numéro de série en mois |
| <a href="https://support.office.com/article/MROUND-function-c299c3b0-15a5-426d-aa4b-d2d5b3baf427" target="_blank">Fonction MROUND</a> | Renvoie un nombre arrondi au dénominateur souhaité |
| <a href="https://support.office.com/article/MULTINOMIAL-function-6fa6373c-6533-41a2-a45e-a56db1db1bf6" target="_blank">Fonction MULTINOMIAL</a> | Calcule la multinomiale d’un ensemble de nombres |
| <a href="https://support.office.com/article/N-function-a624cad1-3635-4208-b54a-29733d1278c9" target="_blank">Fonction N</a> | Renvoie une valeur convertie en nombre |
| <a href="https://support.office.com/article/NA-function-5469c2d1-a90c-4fb5-9bbc-64bd9bb6b47c" target="_blank">Fonction NA</a> | Renvoie la valeur d’erreur #N/A |
| <a href="https://support.office.com/article/NEGBINOMDIST-function-c8239f89-c2d0-45bd-b6af-172e570f8599" target="_blank">Fonction NEGBINOM.DIST</a> | Renvoie la distribution négative binomiale |
| <a href="https://support.office.com/article/NETWORKDAYS-function-48e717bf-a7a3-495f-969e-5005e3eb18e7" target="_blank">Fonction NETWORKDAYS</a> | Renvoie le nombre de jours ouvrés entiers entre deux dates |
| <a href="https://support.office.com/article/NETWORKDAYSINTL-function-a9b26239-4f20-46a1-9ab8-4e925bfd5e28" target="_blank">Fonction NETWORKDAYS.INTL</a> | Renvoie le nombre de jours ouvrés entiers compris entre deux dates à l’aide de paramètres indiquant le nombre de jours compris dans un week-end |
| <a href="https://support.office.com/article/NOMINAL-function-7f1ae29b-6b92-435e-b950-ad8b190ddd2b" target="_blank">Fonction NOMINAL</a> | Renvoie le taux d’intérêt nominal annuel |
| <a href="https://support.office.com/article/NORMDIST-function-edb1cc14-a21c-4e53-839d-8082074c9f8d" target="_blank">Fonction NORM.DIST</a> | Renvoie la distribution cumulée suivant une loi normale |
| <a href="https://support.office.com/article/NORMINV-function-54b30935-fee7-493c-bedb-2278a9db7e13" target="_blank">Fonction NORM.INV</a> | Renvoie l’inverse de la distribution cumulée suivant une loi normale |
| <a href="https://support.office.com/article/NORMSDIST-function-1e787282-3832-4520-a9ae-bd2a8d99ba88" target="_blank">Fonction NORM.S.DIST</a> | Renvoie la distribution cumulée suivant une loi normale centrée réduite |
| <a href="https://support.office.com/article/NORMSINV-function-d6d556b4-ab7f-49cd-b526-5a20918452b1" target="_blank">Fonction NORM.S.INV</a> | Renvoie l’inverse de la distribution cumulée suivant une loi normale centrée réduite |
| <a href="https://support.office.com/article/NOT-function-9cfc6011-a054-40c7-a140-cd4ba2d87d77" target="_blank">Fonction NOT</a> | Inverse la logique de son argument |
| <a href="https://support.office.com/article/NOW-function-3337fd29-145a-4347-b2e6-20c904739c46" target="_blank">Fonction NOW</a> | Renvoie le numéro de série de la date et de l’heure actuelles |
| <a href="https://support.office.com/article/NPER-function-240535b5-6653-4d2d-bfcf-b6a38151d815" target="_blank">Fonction NPER</a> | Renvoie le nombre de paiements d’un investissement |
| <a href="https://support.office.com/article/NPV-function-8672cb67-2576-4d07-b67b-ac28acf2a568" target="_blank">Fonction NPV</a> | Renvoie la valeur nette actuelle d’un investissement, en fonction d’une série de flux de trésorerie périodiques et d’un taux d’escompte |
| <a href="https://support.office.com/article/NUMBERVALUE-function-1b05c8cf-2bfa-4437-af70-596c7ea7d879" target="_blank">Fonction NUMBERVALUE</a> | Convertit le texte en nombre quels que soient les paramètres régionaux |
| <a href="https://support.office.com/article/OCT2BIN-function-55383471-3c56-4d27-9522-1a8ec646c589" target="_blank">Fonction OCT2BIN</a> | Convertit un nombre octal en nombre binaire |
| <a href="https://support.office.com/article/OCT2DEC-function-87606014-cb98-44b2-8dbb-e48f8ced1554" target="_blank">Fonction OCT2DEC</a> | Convertit un nombre octal en nombre décimal |
| <a href="https://support.office.com/article/OCT2HEX-function-912175b4-d497-41b4-a029-221f051b858f" target="_blank">Fonction OCT2HEX</a> | Convertit un nombre octal en nombre hexadécimal |
| <a href="https://support.office.com/article/ODD-function-deae64eb-e08a-4c88-8b40-6d0b42575c98" target="_blank">Fonction ODD</a> | Arrondit un nombre à l’entier impair supérieur le plus proche |
| <a href="https://support.office.com/article/ODDFPRICE-function-d7d664a8-34df-4233-8d2b-922bcf6a69e1" target="_blank">Fonction ODDFPRICE</a> | Renvoie le prix par valeur faciale de 100 dollars d’un titre dont la première période est irrégulière |
| <a href="https://support.office.com/article/ODDFYIELD-function-66bc8b7b-6501-4c93-9ce3-2fd16220fe37" target="_blank">Fonction ODDFYIELD</a> | Renvoie le rendement d’un titre dont la première période est irrégulière |
| <a href="https://support.office.com/article/ODDLPRICE-function-fb657749-d200-4902-afaf-ed5445027fc4" target="_blank">Fonction ODDLPRICE</a> | Renvoie le prix par valeur faciale de 100 dollars d’un titre dont la dernière période est irrégulière |
| <a href="https://support.office.com/article/ODDLYIELD-function-c873d088-cf40-435f-8d41-c8232fee9238" target="_blank">Fonction ODDLYIELD</a> | Renvoie le rendement d’un titre dont la dernière période est irrégulière |
| <a href="https://support.office.com/article/OR-function-7d17ad14-8700-4281-b308-00b131e22af0" target="_blank">Fonction OR</a> | Renvoie `TRUE` si un argument a la valeur True |
| <a href="https://support.office.com/article/PDURATION-function-44f33460-5be5-4c90-b857-22308892adaf" target="_blank">Fonction PDURATION</a> | Renvoie le nombre de périodes requises par un investissement pour atteindre une valeur spécifiée |
| <a href="https://support.office.com/article/PERCENTILEEXC-function-bbaa7204-e9e1-4010-85bf-c31dc5dce4ba" target="_blank">Fonction PERCENTILE.EXC</a> | Renvoie le k-ième centile de valeur d’une plage, où k se trouve dans la plage de 0 à 1 exclus |
| <a href="https://support.office.com/article/PERCENTILEINC-function-680f9539-45eb-410b-9a5e-c1355e5fe2ed" target="_blank">Fonction PERCENTILE.INC</a> | Renvoie le k-ième centile des valeurs d’une plage |
| <a href="https://support.office.com/article/PERCENTRANKEXC-function-d8afee96-b7e2-4a2f-8c01-8fcdedaa6314" target="_blank">Fonction PERCENTRANK.EXC</a> | Renvoie le rang d’une valeur dans un ensemble de données défini comme pourcentage (0..1, exclus) de cet ensemble |
| <a href="https://support.office.com/article/PERCENTRANKINC-function-149592c9-00c0-49ba-86c1-c1f45b80463a" target="_blank">Fonction PERCENTRANK.INC</a> | Renvoie le rang en pourcentage d’une valeur dans un jeu de données |
| <a href="https://support.office.com/article/PERMUT-function-3bd1cb9a-2880-41ab-a197-f246a7a602d3" target="_blank">Fonction PERMUT</a> | Renvoie le nombre de permutations pour un nombre d’objets donné |
| <a href="https://support.office.com/article/PERMUTATIONA-function-6c7d7fdc-d657-44e6-aa19-2857b25cae4e" target="_blank">Fonction PERMUTATIONA</a> | Renvoie le nombre de permutations pour un nombre d’objets donné (avec répétitions) pouvant être sélectionnés à partir du nombre total d’objets |
| <a href="https://support.office.com/article/PHI-function-23e49bc6-a8e8-402d-98d3-9ded87f6295c" target="_blank">Fonction PHI</a> | Renvoie la valeur de la fonction de densité pour une distribution suivant une loi normale centrée réduite |
| <a href="https://support.office.com/article/PI-function-264199d0-a3ba-46b8-975a-c4a04608989b" target="_blank">Fonction PI</a> | Renvoie la valeur de pi |
| <a href="https://support.office.com/article/PMT-function-0214da64-9a63-4996-bc20-214433fa6441" target="_blank">Fonction PMT</a> | Renvoie le montant périodique d’une annuité |
| <a href="https://support.office.com/article/POISSONDIST-function-8fe148ff-39a2-46cb-abf3-7772695d9636" target="_blank">Fonction POISSON.DIST</a> | Renvoie la distribution suivant une loi de Poisson |
| <a href="https://support.office.com/article/POWER-function-d3f2908b-56f4-4c3f-895a-07fb519c362a" target="_blank">Fonction POWER</a> | Renvoie le résultat d’un nombre élevé à une puissance |
| <a href="https://support.office.com/article/PPMT-function-c370d9e3-7749-4ca4-beea-b06c6ac95e1b" target="_blank">Fonction PPMT</a> | Renvoie la part de remboursement du principal d’un emprunt pour une période donnée |
| <a href="https://support.office.com/article/PRICE-function-3ea9deac-8dfa-436f-a7c8-17ea02c21b0a" target="_blank">Fonction PRICE</a> | Renvoie le prix par valeur faciale de 100 dollars d’un titre dont les intérêts sont payés périodiquement |
| <a href="https://support.office.com/article/PRICEDISC-function-d06ad7c1-380e-4be7-9fd9-75e3079acfd3" target="_blank">Fonction PRICEDISC</a> | Renvoie le prix par valeur faciale de 100 dollars pour un titre escompté |
| <a href="https://support.office.com/article/PRICEMAT-function-52c3b4da-bc7e-476a-989f-a95f675cae77" target="_blank">Fonction PRICEMAT</a> | Renvoie le prix par valeur faciale de 100 dollars d’un titre dont les intérêts sont payés à échéance |
| <a href="https://support.office.com/article/PRODUCT-function-8e6b5b24-90ee-4650-aeec-80982a0512ce" target="_blank">PRODUIT, fonction</a> | Multiplie ses arguments |
| <a href="https://support.office.com/article/PROPER-function-52a5a283-e8b2-49be-8506-b2887b889f94" target="_blank">Fonction PROPER</a> | Met en majuscule la première lettre de chaque mot d’une valeur textuelle |
| <a href="https://support.office.com/article/PV-function-23879d31-0e02-4321-be01-da16e8168cbd" target="_blank">Fonction PV</a> | Renvoie la valeur actuelle d’un investissement |
| <a href="https://support.office.com/article/QUARTILEEXC-function-5a355b7a-840b-4a01-b0f1-f538c2864cad" target="_blank">Fonction QUARTILE.EXC</a> | Renvoie le quartile de l’ensemble de données d’après des valeurs de centile comprises entre 0 et 1, exclus |
| <a href="https://support.office.com/article/QUARTILEINC-function-1bbacc80-5075-42f1-aed6-47d735c4819d" target="_blank">Fonction QUARTILE.INC</a> | Renvoie le quartile d’un jeu de données |
| <a href="https://support.office.com/article/QUOTIENT-function-9f7bf099-2a18-4282-8fa4-65290cc99dee" target="_blank">Fonction QUOTIENT</a> | Renvoie la partie entière d’une division |
| <a href="https://support.office.com/article/RADIANS-function-ac409508-3d48-45f5-ac02-1497c92de5bf" target="_blank">Fonction RADIANS</a> | Convertit des degrés en radians |
| <a href="https://support.office.com/article/RAND-function-4cbfa695-8869-4788-8d90-021ea9f5be73" target="_blank">Fonction RAND</a> | Renvoie un nombre aléatoire compris entre 0 et 1 |
| <a href="https://support.office.com/article/RANDBETWEEN-function-4cc7f0d1-87dc-4eb7-987f-a469ab381685" target="_blank">Fonction RANDBETWEEN</a> | Renvoie un nombre aléatoire entre les nombres que vous spécifiez |
| <a href="https://support.office.com/article/RANKAVG-function-bd406a6f-eb38-4d73-aa8e-6d1c3c72e83a" target="_blank">Fonction RANK.AVG</a> | Renvoie le rang d’un nombre dans une liste de nombres |
| <a href="https://support.office.com/article/RANKEQ-function-284858ce-8ef6-450e-b662-26245be04a40" target="_blank">Fonction RANK.EQ</a> | Renvoie le rang d’un nombre dans une liste de nombres |
| <a href="https://support.office.com/article/RATE-function-9f665657-4a7e-4bb7-a030-83fc59e748ce" target="_blank">Fonction RATE</a> | Renvoie le taux d’intérêt par période pour une annuité |
| <a href="https://support.office.com/article/RECEIVED-function-7a3f8b93-6611-4f81-8576-828312c9b5e5" target="_blank">Fonction RECEIVED</a> | Renvoie le montant reçu lorsqu’un titre totalement investi arrive à échéance |
| <a href="https://support.office.com/article/REPLACE-REPLACEB-functions-8d799074-2425-4a8a-84bc-82472868878a" target="_blank">Fonctions REPLACE, REPLACEB</a> | Remplace des caractères dans un texte |
| <a href="https://support.office.com/article/REPT-function-04c4d778-e712-43b4-9c15-d656582bb061" target="_blank">Fonction REPT</a> | Répète un texte un certain nombre de fois |
| <a href="https://support.office.com/article/RIGHT-RIGHTB-functions-240267ee-9afa-4639-a02b-f19e1786cf2f" target="_blank">Fonctions RIGHT, RIGHTB</a> | Renvoie les caractères les plus à droite d’une valeur textuelle |
| <a href="https://support.office.com/article/ROMAN-function-d6b0b99e-de46-4704-a518-b45a0f8b56f5" target="_blank">Fonction ROMAN</a> | Convertit un numéral arabe en caractères romans, sous forme de texte |
| <a href="https://support.office.com/article/ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c" target="_blank">Fonction ROUND</a> | Arrondit un nombre à un nombre de chiffres spécifié |
| <a href="https://support.office.com/article/ROUNDDOWN-function-2ec94c73-241f-4b01-8c6f-17e6d7968f53" target="_blank">Fonction ROUNDDOWN</a> | Arrondit un nombre à la valeur d’arrondi la plus proche de zéro |
| <a href="https://support.office.com/article/ROUNDUP-function-f8bc9b23-e795-47db-8703-db171d0c42a7" target="_blank">Fonction ROUNDUP</a> | Arrondit un nombre à la valeur d’arrondi la plus éloignée de zéro |
| <a href="https://support.office.com/article/ROWS-function-b592593e-3fc2-47f2-bec1-bda493811597" target="_blank">Fonction ROWS</a> | Renvoie le nombre de lignes dans une référence |
| <a href="https://support.office.com/article/RRI-function-6f5822d8-7ef1-4233-944c-79e8172930f4" target="_blank">Fonction RRI</a> | Renvoie un taux d’intérêt équivalent pour la croissance d’un investissement |
| <a href="https://support.office.com/article/SEC-function-ff224717-9c87-4170-9b58-d069ced6d5f7" target="_blank">Fonction SEC</a> | Renvoie la sécante d’un angle |
| <a href="https://support.office.com/article/SECH-function-e05a789f-5ff7-4d7f-984a-5edb9b09556f" target="_blank">Fonction SECH</a> | Renvoie la sécante hyperbolique d’un angle |
| <a href="https://support.office.com/article/SECOND-function-740d1cfc-553c-4099-b668-80eaa24e8af1" target="_blank">Fonction SECOND</a> | Convertit un numéro de série en seconde |
| <a href="https://support.office.com/article/SERIESSUM-function-a3ab25b5-1093-4f5b-b084-96c49087f637" target="_blank">Fonction SERIESSUM</a> | Renvoie le total d’une série de puissance basé sur la formule |
| <a href="https://support.office.com/article/SHEET-function-44718b6f-8b87-47a1-a9d6-b701c06cff24" target="_blank">Fonction SHEET</a> | Renvoie le numéro de la feuille référencée |
| <a href="https://support.office.com/article/SHEETS-function-770515eb-e1e8-45ce-8066-b557e5e4b80b" target="_blank">Fonction SHEETS</a> | Renvoie le nombre de feuilles dans une référence |
| <a href="https://support.office.com/article/SIGN-function-109c932d-fcdc-4023-91f1-2dd0e916a1d8" target="_blank">Fonction SIGN</a> | Renvoie le signe d’un nombre |
| <a href="https://support.office.com/article/SIN-function-cf0e3432-8b9e-483c-bc55-a76651c95602" target="_blank">Fonction SIN</a> | Renvoie le sinus d’un angle donné |
| <a href="https://support.office.com/article/SINH-function-1e4e8b9f-2b65-43fc-ab8a-0a37f4081fa7" target="_blank">Fonction SINH</a> | Renvoie le sinus hyperbolique d’un nombre |
| <a href="https://support.office.com/article/SKEW-function-bdf49d86-b1ef-4804-a046-28eaea69c9fa" target="_blank">Fonction SKEW</a> | Renvoie l’asymétrie d’une distribution |
| <a href="https://support.office.com/article/SKEWP-function-76530a5c-99b9-48a1-8392-26632d542fcb" target="_blank">Fonction SKEW.P</a> | Renvoie l’asymétrie d’une distribution en fonction d’une population : la caractérisation du degré d’asymétrie d’une distribution par rapport à sa moyenne |
| <a href="https://support.office.com/article/SLN-function-cdb666e5-c1c6-40a7-806a-e695edc2f1c8" target="_blank">Fonction SLN</a> | Renvoie l’amortissement linéaire d’une immobilisation pour une période |
| <a href="https://support.office.com/article/SMALL-function-17da8222-7c82-42b2-961b-14c45384df07" target="_blank">Fonction SMALL</a> | Renvoie la k-ième plus petite valeur d’un jeu de données |
| <a href="https://support.office.com/article/SQRT-function-654975c2-05c4-4831-9a24-2c65e4040fdf" target="_blank">Fonction SQRT</a> | Renvoie une racine carrée positive |
| <a href="https://support.office.com/article/SQRTPI-function-1fb4e63f-9b51-46d6-ad68-b3e7a8b519b4" target="_blank">Fonction SQRTPI</a> | Renvoie la racine carrée de (nombre * pi). |
| <a href="https://support.office.com/article/STANDARDIZE-function-81d66554-2d54-40ec-ba83-6437108ee775" target="_blank">Fonction STANDARDIZE</a> | Renvoie une valeur normalisée |
| <a href="https://support.office.com/article/STDEVP-function-6e917c05-31a0-496f-ade7-4f4e7462f285" target="_blank">Fonction STDEV.P</a> | Calcule l’écart type en fonction de la population entière |
| <a href="https://support.office.com/article/STDEVS-function-7d69cf97-0c1f-4acf-be27-f3e83904cc23" target="_blank">Fonction STDEV.S</a> | Évalue l’écart type en fonction d’un échantillon |
| <a href="https://support.office.com/article/STDEVA-function-5ff38888-7ea5-48de-9a6d-11ed73b29e9d" target="_blank">Fonction STDEVA</a> | Évalue l’écart type en fonction d’un échantillon, y compris les nombres, le texte et les valeurs logiques |
| <a href="https://support.office.com/article/STDEVPA-function-5578d4d6-455a-4308-9991-d405afe2c28c" target="_blank">Fonction STDEVPA</a> | Calcule l’écart type en fonction de la population entière, y compris les nombres, le texte et les valeurs logiques |
| <a href="https://support.office.com/article/SUBSTITUTE-function-6434944e-a904-4336-a9b0-1e58df3bc332" target="_blank">Fonction SUBSTITUTE</a> | Remplace le nouveau texte par l’ancien texte d’une chaîne de texte |
| <a href="https://support.office.com/article/SUBTOTAL-function-7b027003-f060-4ade-9040-e478765b9939" target="_blank">Fonction SUBTOTAL</a> | Renvoie un sous-total dans une liste ou une base de données |
| <a href="https://support.office.com/article/SUM-function-043e1c7d-7726-4e80-8f32-07b23e057f89" target="_blank">Fonction SUM</a> | Ajoute ses arguments |
| <a href="https://support.office.com/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b" target="_blank">Fonction SUMIF</a> | Ajoute les cellules spécifiées par un critère donné |
| <a href="https://support.office.com/article/SUMIFS-function-c9e748f5-7ea7-455d-9406-611cebce642b" target="_blank">Fonction SUMIFS</a> | Ajoute les cellules d’une plage répondant à plusieurs critères |
| <a href="https://support.office.com/article/SUMSQ-function-e3313c02-51cc-4963-aae6-31442d9ec307" target="_blank">Fonction SUMSQ</a> | Renvoie le total des carrés des arguments |
| <a href="https://support.office.com/article/SYD-function-069f8106-b60b-4ca2-98e0-2a0f206bdb27" target="_blank">Fonction SYD</a> | Renvoie l’amortissement des chiffres cumulés sur l’année d’une immobilisation pour une période spécifique |
| <a href="https://support.office.com/article/T-function-fb83aeec-45e7-4924-af95-53e073541228" target="_blank">Fonction T</a> | Convertit ses arguments en texte |
| <a href="https://support.office.com/article/TDIST-function-4329459f-ae91-48c2-bba8-1ead1c6c21b2" target="_blank">Fonction T.DIST</a> | Renvoie les points de pourcentage (probabilité) pour la distribution suivant la loi T de Student |
| <a href="https://support.office.com/article/TDIST2T-function-198e9340-e360-4230-bd21-f52f22ff5c28" target="_blank">Fonction T.DIST.2T</a> | Renvoie les points de pourcentage (probabilité) pour la distribution suivant la loi T de Student |
| <a href="https://support.office.com/article/TDISTRT-function-20a30020-86f9-4b35-af1f-7ef6ae683eda" target="_blank">Fonction T.DIST.RT</a> | Renvoie la distribution suivant la loi T de Student |
| <a href="https://support.office.com/article/TINV-function-2908272b-4e61-4942-9df9-a25fec9b0e2e" target="_blank">Fonction T.INV</a> | Renvoie la valeur t de la distribution suivant la loi T de Student sous forme de fonction de probabilité et de degrés de liberté |
| <a href="https://support.office.com/article/TINV2T-function-ce72ea19-ec6c-4be7-bed2-b9baf2264f17" target="_blank">Fonction T.INV.2T</a> | Renvoie l’inverse de la distribution suivant la loi T de Student |
| <a href="https://support.office.com/article/TAN-function-08851a40-179f-4052-b789-d7f699447401" target="_blank">Fonction TAN</a> | Renvoie la tangente d’un nombre |
| <a href="https://support.office.com/article/TANH-function-017222f0-a0c3-4f69-9787-b3202295dc6c" target="_blank">Fonction TANH</a> | Renvoie la tangente hyperbolique d’un nombre |
| <a href="https://support.office.com/article/TBILLEQ-function-2ab72d90-9b4d-4efe-9fc2-0f81f2c19c8c" target="_blank">Fonction TBILLEQ</a> | Renvoie le rapport lié aux titres pour un bon du Trésor |
| <a href="https://support.office.com/article/TBILLPRICE-function-eacca992-c29d-425a-9eb8-0513fe6035a2" target="_blank">Fonction TBILLPRICE</a> | Renvoie le prix par valeur faciale de 100 dollars pour un bon du Trésor |
| <a href="https://support.office.com/article/TBILLYIELD-function-6d381232-f4b0-4cd5-8e97-45b9c03468ba" target="_blank">Fonction TBILLYIELD</a> | Renvoie le rapport pour un bon du Trésor |
| <a href="https://support.office.com/article/TEXT-function-20d5ac4d-7b94-49fd-bb38-93d29371225c" target="_blank">Fonction TEXT</a> | Met en forme un nombre et le convertit en texte |
| <a href="https://support.office.com/article/TIME-function-9a5aff99-8f7d-4611-845e-747d0b8d5457" target="_blank">Fonction TIME</a> | Renvoie le numéro de série d’une heure précise |
| <a href="https://support.office.com/article/TIMEVALUE-function-0b615c12-33d8-4431-bf3d-f3eb6d186645" target="_blank">Fonction TIMEVALUE</a> | Convertit une heure au format texte en numéro de série |
| <a href="https://support.office.com/article/TODAY-function-5eb3078d-a82c-4736-8930-2f51a028fdd9" target="_blank">Fonction TODAY</a> | Renvoie le numéro de série de la date du jour |
| <a href="https://support.office.com/article/TRIM-function-410388fa-c5df-49c6-b16c-9e5630b479f9" target="_blank">Fonction TRIM</a> | Supprime les espaces du texte |
| <a href="https://support.office.com/article/TRIMMEAN-function-d90c9878-a119-4746-88fa-63d988f511d3" target="_blank">Fonction TRIMMEAN</a> | Renvoie la moyenne de la partie intérieure d’un jeu de données |
| <a href="https://support.office.com/article/TRUE-function-7652c6e3-8987-48d0-97cd-ef223246b3fb" target="_blank">Fonction TRUE</a> | Renvoie la valeur logique `TRUE` |
| <a href="https://support.office.com/article/TRUNC-function-8b86a64c-3127-43db-ba14-aa5ceb292721" target="_blank">Fonction TRUNC</a> | Tronque un nombre en entier |
| <a href="https://support.office.com/article/TYPE-function-45b4e688-4bc3-48b3-a105-ffa892995899" target="_blank">Fonction TYPE</a> | Renvoie un nombre indiquant le type de données d’une valeur |
| <a href="https://support.office.com/article/UNICHAR-function-ffeb64f5-f131-44c6-b332-5cd72f0659b8" target="_blank">Fonction UNICHAR</a> | Renvoie le caractère unicode référencé par la valeur numérique donnée |
| <a href="https://support.office.com/article/UNICODE-function-adb74aaa-a2a5-4dde-aff6-966e4e81f16f" target="_blank">Fonction UNICODE</a> | Renvoie le nombre (point de code) qui correspond au premier caractère du texte |
| <a href="https://support.office.com/article/UPPER-function-c11f29b3-d1a3-4537-8df6-04d0049963d6" target="_blank">Fonction UPPER</a> | Convertit le texte en majuscules |
| <a href="https://support.office.com/article/VALUE-function-257d0108-07dc-437d-ae1c-bc2d3953d8c2" target="_blank">Fonction VALUE</a> | Convertit un argument textuel en nombre |
| <a href="https://support.office.com/article/VARP-function-73d1285c-108c-4843-ba5d-a51f90656f3a" target="_blank">Fonction VAR.P</a> | Calcule l’écart en fonction de la population entière |
| <a href="https://support.office.com/article/VARS-function-913633de-136b-449d-813e-65a00b2b990b" target="_blank">Fonction VAR.S</a> | Fournit une estimation de l’écart à partir d’un échantillon |
| <a href="https://support.office.com/article/VARA-function-3de77469-fa3a-47b4-85fd-81758a1e1d07" target="_blank">Fonction VARA</a> | Évalue la varianceen fonction d’un échantillon, y compris les nombres, le texte et les valeurs logiques |
| <a href="https://support.office.com/article/VARPA-function-59a62635-4e89-4fad-88ac-ce4dc0513b96" target="_blank">Fonction VARPA</a> | Calcule la variance en fonction de la population entière, y compris les nombres, le texte et les valeurs logiques |
| <a href="https://support.office.com/article/VDB-function-dde4e207-f3fa-488d-91d2-66d55e861d73" target="_blank">Fonction VDB</a> | Renvoie l’amortissement d’un bien durant une période spécifiée ou partielle en utilisant une méthode d’amortissement dégressif |
| <a href="https://support.office.com/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1" target="_blank">Fonction VLOOKUP</a> | Cherche dans la première colonne d’un tableau et se déplace horizontalement pour renvoyer la valeur d’une cellule |
| <a href="https://support.office.com/article/WEEKDAY-function-60e44483-2ed1-439f-8bd0-e404c190949a" target="_blank">Fonction WEEKDAY</a> | Convertit un numéro de série en jour de la semaine |
| <a href="https://support.office.com/article/WEEKNUM-function-e5c43a03-b4ab-426c-b411-b18c13c75340" target="_blank">Fonction WEEKNUM</a> | Convertit un numéro de série en un numéro de semaine correspondant à l’année |
| <a href="https://support.office.com/article/WEIBULLDIST-function-4e783c39-9325-49be-bbc9-a83ef82b45db" target="_blank">Fonction WEIBULL.DIST</a> | Renvoie la distribution suivant la loi de Weibull |
| <a href="https://support.office.com/article/WORKDAY-function-f764a5b7-05fc-4494-9486-60d494efbf33" target="_blank">Fonction WORKDAY</a> | Renvoie le numéro de série de la date précédant ou suivant un nombre de jours ouvrés spécifié |
| <a href="https://support.office.com/article/WORKDAYINTL-function-a378391c-9ba7-4678-8a39-39611a9bf81d" target="_blank">Fonction WORKDAY.INTL</a> | Renvoie le numéro de série de la date précédant ou suivant un nombre spécifié de jours ouvrés à l’aide de paramètres indiquant le nombre de jours compris dans un week-end |
| <a href="https://support.office.com/article/XIRR-function-de1242ec-6477-445b-b11b-a303ad9adc9d" target="_blank">Fonction XIRR</a> | Renvoie le taux de rendement interne d’une planification de flux financiers qui n’est pas nécessairement périodique |
| <a href="https://support.office.com/article/XNPV-function-1b42bbf6-370f-4532-a0eb-d67c16b664b7" target="_blank">Fonction XNPV</a> | Renvoie la valeur actuelle nette d’une planification de flux financiers qui n’est pas nécessairement périodique |
| <a href="https://support.office.com/article/XOR-function-1548d4c2-5e47-4f77-9a92-0533bba14f37" target="_blank">Fonction XOR</a> | Renvoie une valeur logique exclusive OR de tous les arguments |
| <a href="https://support.office.com/article/YEAR-function-c64f017a-1354-490d-981f-578e8ec8d3b9" target="_blank">Fonction YEAR</a> | Convertit un numéro de série en année |
| <a href="https://support.office.com/article/YEARFRAC-function-3844141e-c76d-4143-82b6-208454ddc6a8" target="_blank">Fonction YEARFRAC</a> | Renvoie la fraction de l’année représentant le nombre de jours entiers compris entre start_date et end_date |
| <a href="https://support.office.com/article/YIELD-function-f5f5ca43-c4bd-434f-8bd2-ed3c9727a4fe" target="_blank">Fonction YIELD</a> | Renvoie le rendement d’un titre rapportant des intérêts périodiquement |
| <a href="https://support.office.com/article/YIELDDISC-function-a9dbdbae-7dae-46de-b995-615faffaaed7" target="_blank">Fonction YIELDDISC</a> | Renvoie le rendement annuel d’un titre escompté, par exemple, un bon du Trésor |
| <a href="https://support.office.com/article/YIELDMAT-function-ba7d1809-0d33-4bcb-96c7-6c56ec62ef6f" target="_blank">Fonction YIELDMAT</a> | Renvoie le rendement annuel d’un titre pour lequel des intérêts sont payés à l’échéance |
| <a href="https://support.office.com/article/ZTEST-function-d633d5a3-2031-4614-a016-92180ad82bee" target="_blank">Fonction Z.TEST</a> | Renvoie la valeur de probabilité unilatérale du test Z |

## <a name="see-also"></a>Voir aussi

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Classe de fonctions (interface API JavaScript pour Excel)](/javascript/api/excel/excel.functions)
- [Objet de fonctions Workbook (interface API JavaScript pour Excel)](/javascript/api/excel/excel.workbook#functions)
