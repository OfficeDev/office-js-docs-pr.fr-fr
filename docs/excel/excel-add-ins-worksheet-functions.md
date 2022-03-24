---
title: Appel de fonctions de feuille de calcul Excel intégrées à l’aide de l’API JavaScript pour Excel
description: Découvrez comment appeler des fonctions de feuille de Excel intégrées, telles `VLOOKUP` `SUM` que l’API JavaScript Excel et à l’aide de cette fonction.
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: cc7622b642720a8cb8f80ad553600fd22ac7c25c
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744074"
---
# <a name="call-built-in-excel-worksheet-functions"></a>Appel de fonctions de feuille de calcul Excel intégrées

Cet article explique comment appeler les fonctions de feuille de calcul Excel intégrées telles que `VLOOKUP` et `SUM` utilisant l’API JavaScript pour Excel. Il fournit également la liste complète des fonctions de feuille de calcul Excel intégrées pouvant être appelées à l’aide de l’API JavaScript pour Excel.

> [!NOTE]
> Pour plus d’informations sur la création de *fonctions personnalisées* dans Excel à l’aide de l’API JavaScript pour Excel, reportez-vous à la rubrique [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md).

## <a name="calling-a-worksheet-function"></a>Appel d’une fonction de feuille de calcul

L’extrait de code suivant montre comment appeler une fonction de feuille de calcul où `sampleFunction()` est un espace réservé devant être remplacé par le nom de la fonction à appeler et les paramètres d’entrée nécessitant la fonction. La `value` propriété de l’objet `FunctionResult` renvoyé par une fonction de feuille de calcul contient le résultat de la fonction spécifiée. Comme le montre cet exemple, vous devez avoir `load` la propriété `value` de l’objet `FunctionResult` avant de pouvoir le lire. Dans cet exemple, le résultat de la fonction est simplement écrit sur la console.

```js
await Excel.run(async (context) => {
    let functionResult = context.workbook.functions.sampleFunction();
    functionResult.load('value');

    await context.sync();
    console.log('Result of the function: ' + functionResult.value);
});
```

> [!TIP]
> Reportez-vous à la section [Fonctions de feuille de calcul prises en charge](#supported-worksheet-functions) de cet article pour obtenir la liste des fonctions appelées à l’aide de l’API JavaScript pour Excel.

## <a name="sample-data"></a>Exemple de données

L’image suivante montre un tableau dans une feuille de calcul Excel contenant des données de ventes pour divers types d’outils sur une période de trois mois. Chaque numéro de la table représente le nombre d’unités vendues pour un outil spécifique lors d’un mois donné. Les exemples suivant expliquent comment appliquer des fonctions de feuille de calcul intégrées à ces données.

![Capture d’écran des données de ventes Excel données pour hammer, Wrench et Saw en novembre, décembre et janvier.](../images/worksheet-functions-chaining-results.jpg)

## <a name="example-1-single-function"></a>Exemple 1 : Fonction unique

L’exemple de code suivant applique la fonction `VLOOKUP` aux exemples de données décrits précédemment pour identifier le nombre de clés vendues au mois de novembre.

```js
await Excel.run(async (context) => {
    let range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    let unitSoldInNov = context.workbook.functions.vlookup("Wrench", range, 2, false);
    unitSoldInNov.load('value');

    await context.sync();
    console.log(' Number of wrenches sold in November = ' + unitSoldInNov.value);
});
```

## <a name="example-2-nested-functions"></a>Exemple 2 : Fonctions imbriquées

L’exemple de code suivant applique la fonction `VLOOKUP` pour les exemples de données décrits précédemment afin d’identifier le nombre de clés vendues au mois de novembre et le nombre de clés vendues en décembre, puis applique la fonction `SUM` pour calculer le nombre total de clés vendues au cours de ces deux mois.

Comme indiqué dans cet exemple, si un ou plusieurs appels de fonction sont imbriqués dans un autre appel de fonction, vous devez uniquement charger (`load`) le résultat final que vous souhaitez lire par la suite (dans cet exemple, `sumOfTwoLookups`). Les résultats intermédiaires (dans cet exemple, le résultat de chaque fonction `VLOOKUP`) sont calculés et utilisés pour calculer le résultat final.

```js
await Excel.run(async (context) => {
    let range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    let sumOfTwoLookups = context.workbook.functions.sum(
        context.workbook.functions.vlookup("Wrench", range, 2, false),
        context.workbook.functions.vlookup("Wrench", range, 3, false)
    );
    sumOfTwoLookups.load('value');

    await context.sync();
    console.log(' Number of wrenches sold in November and December = ' + sumOfTwoLookups.value);
});
```

## <a name="supported-worksheet-functions"></a>Fonctions de feuille de calcul prises en charge

Les fonctions de feuille de calcul Excel intégrées suivantes peuvent être appelées à l’aide de l’API JavaScript pour Excel.

| Fonction | Description |
|:---------------|:-----------|
| [Fonction ABS](https://support.microsoft.com/office/3420200f-5628-4e8c-99da-c99d7c87713c) | Renvoie la valeur absolue d’un nombre |
| [Fonction ACCRINT](https://support.microsoft.com/office/fe45d089-6722-4fb3-9379-e1f911d8dc74) | Renvoie l’intérêt couru non échu d’un titre dont l’intérêt est perçu périodiquement |
| [Fonction ACCRINTM](https://support.microsoft.com/office/f62f01f9-5754-4cc4-805b-0e70199328a7) | Renvoie l’intérêt couru non échu d’un titre dont l’intérêt est perçu à l’échéance |
| [Fonction ACOS](https://support.microsoft.com/office/cb73173f-d089-4582-afa1-76e5524b5d5b) | Renvoie l’arccosinus d’un nombre |
| [Fonction ACOSH](https://support.microsoft.com/office/e3992cc1-103f-4e72-9f04-624b9ef5ebfe) | Renvoie le cosinus hyperbolique inverse d’un nombre |
| [Fonction ACOT](https://support.microsoft.com/office/dc7e5008-fe6b-402e-bdd6-2eea8383d905) | Renvoie l’arccotangente d’un nombre |
| [Fonction ACOTH](https://support.microsoft.com/office/cc49480f-f684-4171-9fc5-73e4e852300f) | Renvoie l’arccotangente hyperbolique d’un nombre |
| [Fonction AMORDEGRC](https://support.microsoft.com/office/a14d0ca1-64a4-42eb-9b3d-b0dededf9e51) | Renvoie l’amortissement correspondant à chaque période comptable en utilisant un coefficient d’amortissement |
| [Fonction AMORLINC](https://support.microsoft.com/office/7d417b45-f7f5-4dba-a0a5-3451a81079a8) | Renvoie l’amortissement correspondant à chaque période comptable |
| [Fonction AND](https://support.microsoft.com/office/5f19b2e8-e1df-4408-897a-ce285a19e9d9) | Renvoie `TRUE` si tous les arguments ont la valeur True |
| [Fonction ARABIC](https://support.microsoft.com/office/9a8da418-c17b-4ef9-a657-9370a30a674f) | Convertit un nombre romain en chiffre arabe |
| [Fonction AREAS](https://support.microsoft.com/office/8392ba32-7a41-43b3-96b0-3695d2ec6152) | Renvoie le nombre de zones dans une référence |
| [Fonction ASC](https://support.microsoft.com/office/0b6abf1c-c663-4004-a964-ebc00b723266) | Convertit les caractères anglais pleine chasse (codés sur deux octets) ou katakana dans une chaîne de caractères en caractères à demi-chasse (codés sur un octet) |
| [Fonction ASIN](https://support.microsoft.com/office/81fb95e5-6d6f-48c4-bc45-58f955c6d347) | Renvoie l’arcsinus d’un nombre |
| [Fonction ASINH](https://support.microsoft.com/office/4e00475a-067a-43cf-926a-765b0249717c) | Renvoie le sinus hyperbolique inverse d’un nombre |
| [Fonction ATAN](https://support.microsoft.com/office/50746fa8-630a-406b-81d0-4a2aed395543) | Renvoie l’arctangente d’un nombre |
| [Fonction ATAN2](https://support.microsoft.com/office/c04592ab-b9e3-4908-b428-c96b3a565033) | Renvoie l’arctangente des coordonnées x et y |
| [Fonction ATANH](https://support.microsoft.com/office/3cd65768-0de7-4f1d-b312-d01c8c930d90) | Renvoie la tangente hyperbolique inverse d’un nombre |
| [Fonction AVEDEV](https://support.microsoft.com/office/58fe8d65-2a84-4dc7-8052-f3f87b5c6639) | Renvoie la moyenne des écarts absolus des points de données par rapport à leur moyenne |
| [Fonction AVERAGE](https://support.microsoft.com/office/047bac88-d466-426c-a32b-8f33eb960cf6) | Renvoie la moyenne de ses arguments |
| [Fonction AVERAGEA](https://support.microsoft.com/office/f5f84098-d453-4f4c-bbba-3d2c66356091) | Renvoie la moyenne de ses arguments, y compris les nombres, le texte et les valeurs logiques |
| [Fonction AVERAGEIF](https://support.microsoft.com/office/faec8e2e-0dec-4308-af69-f5576d8ac642) | Renvoie la moyenne (arithmétique) de toutes les cellules d’une plage respectant un critère donné |
| [Fonction AVERAGEIFS](https://support.microsoft.com/office/48910c45-1fc0-4389-a028-f7c5c3001690) | Renvoie la moyenne (arithmétique) de toutes les cellules qui répondent à plusieurs critères |
| [Fonction BAHTTEXT](https://support.microsoft.com/office/5ba4d0b4-abd3-4325-8d22-7a92d59aab9c) | Convertit un nombre en texte en utilisant le format monétaire ß (baht) |
| [Fonction BASE](https://support.microsoft.com/office/2ef61411-aee9-4f29-a811-1c42456c6342) | Convertit un nombre en représentation textuelle avec la base spécifiée |
| [Fonction BESSELI](https://support.microsoft.com/office/8d33855c-9a8d-444b-98e0-852267b1c0df) | Renvoie la fonction de Bessel modifiée In(x) |
| [Fonction BESSELJ](https://support.microsoft.com/office/839cb181-48de-408b-9d80-bd02982d94f7) | Renvoie la fonction de Bessel Jn(x) |
| [Fonction BESSELK](https://support.microsoft.com/office/606d11bc-06d3-4d53-9ecb-2803e2b90b70) | Renvoie la fonction de Bessel modifiée Kn(x) |
| [Fonction BESSELY](https://support.microsoft.com/office/f3a356b3-da89-42c3-8974-2da54d6353a2) | Renvoie la fonction de Bessel Yn(x) |
| [Fonction BETA.DIST](https://support.microsoft.com/office/11188c9c-780a-42c7-ba43-9ecb5a878d31) | Renvoie la fonction de distribution cumulée suivant une loi Bêta |
| [Fonction BETA.INV](https://support.microsoft.com/office/e84cb8aa-8df0-4cf6-9892-83a341d252eb) | Renvoie l’inverse de la fonction de distribution cumulée pour une distribution bêta spécifiée |
| [Fonction BIN2DEC](https://support.microsoft.com/office/63905b57-b3a0-453d-99f4-647bb519cd6c) | Convertit un nombre binaire en nombre décimal |
| [Fonction BIN2HEX](https://support.microsoft.com/office/0375e507-f5e5-4077-9af8-28d84f9f41cc) | Convertit un nombre binaire en nombre hexadécimal |
| [Fonction BIN2OCT](https://support.microsoft.com/office/0a4e01ba-ac8d-4158-9b29-16c25c4c23fd) | Convertit un nombre binaire en nombre octal |
| [Fonction BINOM.DIST](https://support.microsoft.com/office/c5ae37b6-f39c-4be2-94c2-509a1480770c) | Renvoie la probabilité d’une variable aléatoire discrète suivant la loi binomiale |
| [Fonction BINOM.DIST.RANGE](https://support.microsoft.com/office/17331329-74c7-4053-bb4c-6653a7421595) | Renvoie la probabilité d’un résultat de tirage en suivant une distribution binomiale |
| [Fonction BINOM.INV](https://support.microsoft.com/office/80a0370c-ada6-49b4-83e7-05a91ba77ac9) | Renvoie la plus petite valeur pour laquelle la distribution binomiale cumulée est inférieure ou égale à une valeur critère |
| [Fonction BITAND](https://support.microsoft.com/office/8a2be3d7-91c3-4b48-9517-64548008563a) | Renvoie une opération AND au niveau du bit de deux nombres |
| [Fonction BITLSHIFT](https://support.microsoft.com/office/c55bb27e-cacd-4c7c-b258-d80861a03c9c) | Renvoie un nombre décalé vers la gauche de total_décalage bits. |
| [Fonction BITOR](https://support.microsoft.com/office/f6ead5c8-5b98-4c9e-9053-8ad5234919b2) | Renvoie une opération OR au niveau du bit de deux nombres |
| [Fonction BITRSHIFT](https://support.microsoft.com/office/274d6996-f42c-4743-abdb-4ff95351222c) | Renvoie un nombre décalé vers la droite de total_décalage bits |
| [Fonction BITXOR](https://support.microsoft.com/office/c81306a1-03f9-4e89-85ac-b86c3cba10e4) | Renvoie une opération Exclusive Or au niveau du bit de deux nombres |
| [CEILING. MATH, ECMA_CEILING fonctions](https://support.microsoft.com/office/80f95d2f-b499-4eee-9f16-f795a8e306c8) | Arrondit un nombre à l’entier ou au multiple supérieur le plus proche de l’argument de précision |
| [Fonction CEILING.PRECISE](https://support.microsoft.com/office/f366a774-527a-4c92-ba49-af0a196e66cb) | Arrondit un nombre à l’entier ou au multiple le plus proche de l’argument de précision. Quel que soit le signe du nombre, le nombre est arrondi à l’unité supérieure. |
| [Fonction CHAR](https://support.microsoft.com/office/bbd249c8-b36e-4a91-8017-1c133f9b837a) | Renvoie le caractère spécifié par le code numérique |
| [Fonction CHISQ.DIST](https://support.microsoft.com/office/8486b05e-5c05-4942-a9ea-f6b341518732) | Renvoie la fonction de densité de probabilité bêta cumulative |
| [Fonction CHISQ.DIST.RT](https://support.microsoft.com/office/dc4832e8-ed2b-49ae-8d7c-b28d5804c0f2) | Renvoie la probabilité d’une variable aléatoire continue suivant une loi unilatérale du Khi-deux |
| [Fonction CHISQ.INV](https://support.microsoft.com/office/400db556-62b3-472d-80b3-254723e7092f) | Renvoie la fonction de densité de probabilité bêta cumulative |
| [Fonction CHISQ.INV.RT](https://support.microsoft.com/office/435b5ed8-98d5-4da6-823f-293e2cbc94fe) | Renvoie l’inverse de la probabilité d’une variable aléatoire continue suivant une loi unilatérale du Khi-deux |
| [Fonction CHOOSE](https://support.microsoft.com/office/fc5c184f-cb62-4ec7-a46e-38653b98f5bc) | Choisit une valeur dans une liste de valeurs |
| [Fonction CLEAN](https://support.microsoft.com/office/26f3d7c5-475f-4a9c-90e5-4b8ba987ba41) | Supprime tous les caractères non imprimables du texte |
| [Fonction CODE](https://support.microsoft.com/office/c32b692b-2ed0-4a04-bdd9-75640144b928) | Renvoie le code numérique du premier caractère d’une chaîne de texte |
| [Fonction COLUMNS](https://support.microsoft.com/office/4e8e7b4e-e603-43e8-b177-956088fa48ca) | Renvoie le nombre de colonnes dans une référence |
| [Fonction COMBIN](https://support.microsoft.com/office/12a3f276-0a21-423a-8de6-06990aaf638a) | Renvoie le nombre de combinaisons pour un nombre d’objets donné |
| [Fonction COMBINA](https://support.microsoft.com/office/efb49eaa-4f4c-4cd2-8179-0ddfcf9d035d) | Renvoie le nombre de combinaisons avec répétitions pour un nombre d’éléments donné |
| [Fonction COMPLEX](https://support.microsoft.com/office/f0b8f3a9-51cc-4d6d-86fb-3a9362fa4128) | Convertit des coefficients réels et imaginaires en nombre complexe |
| [Fonction CONCATENATE](https://support.microsoft.com/office/8f8ae884-2ca8-4f7a-b093-75d702bea31d) | Regroupe plusieurs éléments textuels en un élément textuel |
| [Fonction CONFIDENCE.NORM](https://support.microsoft.com/office/7cec58a6-85bb-488d-91c3-63828d4fbfd4) | Renvoie l’intervalle de confiance pour la moyenne d’une population |
| [Fonction CONFIDENCE.T](https://support.microsoft.com/office/e8eca395-6c3a-4ba9-9003-79ccc61d3c53) | Renvoie l’intervalle de confiance pour la moyenne d’une population, à l’aide de la probabilité d’une variable aléatoire suivant une loi T de Student |
| [Fonction CONVERT](https://support.microsoft.com/office/d785bef1-808e-4aac-bdcd-666c810f9af2) | Convertit un nombre d’un système de mesure à un autre |
| [Fonction COS](https://support.microsoft.com/office/0fb808a5-95d6-4553-8148-22aebdce5f05) | Renvoie le cosinus d’un nombre |
| [Fonction COSH](https://support.microsoft.com/office/e460d426-c471-43e8-9540-a57ff3b70555) | Renvoie le cosinus hyperbolique d’un nombre |
| [Fonction COT](https://support.microsoft.com/office/c446f34d-6fe4-40dc-84f8-cf59e5f5e31a) | Renvoie la cotangente d’un angle |
| [Fonction COTH](https://support.microsoft.com/office/2e0b4cb6-0ba0-403e-aed4-deaa71b49df5) | Renvoie la cotangente hyperbolique d’un nombre |
| [Fonction COUNT](https://support.microsoft.com/office/a59cd7fc-b623-4d93-87a4-d23bf411294c) | Compte le nombre de chiffres compris dans la liste d’arguments |
| [Fonction COUNTA](https://support.microsoft.com/office/7dc98875-d5c1-46f1-9a82-53f3219e2509) | Compte le nombre de valeurs comprises dans la liste d’arguments |
| [Fonction COUNTBLANK](https://support.microsoft.com/office/6a92d772-675c-4bee-b346-24af6bd3ac22) | Compte le nombre de cellules vides dans une plage |
| [Fonction COUNTIF](https://support.microsoft.com/office/e0de10c6-f885-4e71-abb4-1f464816df34) | Compte le nombre de cellules à l’intérieur d’une plage qui répondent aux critères donnés |
| [Fonction COUNTIFS](https://support.microsoft.com/office/dda3dc6e-f74e-4aee-88bc-aa8c2a866842) | Compte le nombre de cellules à l’intérieur d’une plage qui répondent à plusieurs critères |
| [Fonction COUPDAYBS](https://support.microsoft.com/office/eb9a8dfb-2fb2-4c61-8e5d-690b320cf872) | Renvoie le nombre de jours entre le début de la période du coupon et la date d’escompte |
| [Fonction COUPDAYS](https://support.microsoft.com/office/cc64380b-315b-4e7b-950c-b30b0a76f671) | Renvoie le nombre de jours dans la période du coupon contenant la date d’escompte |
| [Fonction COUPDAYSNC](https://support.microsoft.com/office/5ab3f0b2-029f-4a8b-bb65-47d525eea547) | Renvoie le nombre de jours séparant la date d’escompte de la date du prochain coupon |
| [Fonction COUPNCD](https://support.microsoft.com/office/fd962fef-506b-4d9d-8590-16df5393691f) | Renvoie la date du prochain coupon suivant la date d’escompte |
| [Fonction COUPNUM](https://support.microsoft.com/office/a90af57b-de53-4969-9c99-dd6139db2522) | Renvoie le nombre de coupons à régler entre la date d’escompte et la date d’échéance |
| [Fonction COUPPCD](https://support.microsoft.com/office/2eb50473-6ee9-4052-a206-77a9a385d5b3) | Renvoie la date du coupon antérieur précédant la date d’escompte |
| [Fonction CSC](https://support.microsoft.com/office/07379361-219a-4398-8675-07ddc4f135c1) | Renvoie la cosécante d’un angle |
| [Fonction CSCH](https://support.microsoft.com/office/f58f2c22-eb75-4dd6-84f4-a503527f8eeb) | Renvoie la cosécante hyperbolique d’un angle |
| [Fonction CUMIPMT](https://support.microsoft.com/office/61067bb0-9016-427d-b95b-1a752af0e606) | Renvoie les intérêts cumulés réglés entre deux périodes |
| [Fonction CUMPRINC](https://support.microsoft.com/office/94a4516d-bd65-41a1-bc16-053a6af4c04d) | Renvoie le montant cumulé du remboursement du capital réglé entre deux périodes |
| [Fonction DATE](https://support.microsoft.com/office/e36c0c8c-4104-49da-ab83-82328b832349) | Renvoie le numéro de série d’une date précise |
| [Fonction DATEVALUE](https://support.microsoft.com/office/df8b07d4-7761-4a93-bc33-b7471bbff252) | Convertit une date au format texte en numéro de série |
| [Fonction DAVERAGE](https://support.microsoft.com/office/a6a2d5ac-4b4b-48cd-a1d8-7b37834e5aee) | Renvoie la moyenne des entrées d’une base de données sélectionnée |
| [Fonction DAY](https://support.microsoft.com/office/8a7d1cbb-6c7d-4ba1-8aea-25c134d03101) | Convertit un numéro de série en jour du mois |
| [Fonction DAYS](https://support.microsoft.com/office/57740535-d549-4395-8728-0f07bff0b9df) | Renvoie le nombre de jours entre deux dates |
| [Fonction DAYS360](https://support.microsoft.com/office/b9a509fd-49ef-407e-94df-0cbda5718c2a) | Calcule le nombre de jours entre deux dates sur la base d’une année de 360 jours |
| [Fonction DB](https://support.microsoft.com/office/354e7d28-5f93-4ff1-8a52-eb4ee549d9d7) | Renvoie l’amortissement d’un bien durant une période spécifiée en utilisant la méthode de l’amortissement dégressif à taux fixe |
| [Fonction DBCS](https://support.microsoft.com/office/a4025e73-63d2-4958-9423-21a24794c9e5) | Convertit les caractères anglais à demi-chasse (codés sur un octet) ou katakana dans une chaîne de caractères en caractères pleine chasse (codés sur deux octets) |
| [Fonction DCOUNT](https://support.microsoft.com/office/c1fc7b93-fb0d-4d8d-97db-8d5f076eaeb1) | Compte les cellules qui contiennent des nombres dans une base de données |
| [Fonction DCOUNTA](https://support.microsoft.com/office/00232a6d-5a66-4a01-a25b-c1653fda1244) | Compte les cellules non vides d’une base de données |
| [Fonction DDB](https://support.microsoft.com/office/519a7a37-8772-4c96-85c0-ed2c209717a5) | Renvoie l’amortissement d’un bien durant une période spécifiée suivant la méthode de l’amortissement dégressif à taux double ou selon un coefficient à spécifier |
| [Fonction DEC2BIN](https://support.microsoft.com/office/0f63dd0e-5d1a-42d8-b511-5bf5c6d43838) | Convertit un nombre décimal en nombre binaire |
| [Fonction DEC2HEX](https://support.microsoft.com/office/6344ee8b-b6b5-4c6a-a672-f64666704619) | Convertit un nombre décimal en nombre hexadécimal |
| [Fonction DEC2OCT](https://support.microsoft.com/office/c9d835ca-20b7-40c4-8a9e-d3be351ce00f) | Convertit un nombre décimal en nombre octal |
| [Fonction DECIMAL](https://support.microsoft.com/office/ee554665-6176-46ef-82de-0a283658da2e) | Convertit une représentation textuelle d’un nombre dans une base donnée en nombre décimal |
| [Fonction DEGREES](https://support.microsoft.com/office/4d6ec4db-e694-4b94-ace0-1cc3f61f9ba1) | Convertit des radians en degrés |
| [Fonction DELTA](https://support.microsoft.com/office/2f763672-c959-4e07-ac33-fe03220ba432) | Vérifie si deux valeurs sont égales |
| [Fonction DEVSQ](https://support.microsoft.com/office/8b739616-8376-4df5-8bd0-cfe0a6caf444) | Renvoie la somme des carrés des écarts |
| [Fonction DGET](https://support.microsoft.com/office/455568bf-4eef-45f7-90f0-ec250d00892e) | Extrait d’une base de données un seul enregistrement correspondant aux critères spécifiés |
| [Fonction DISC](https://support.microsoft.com/office/71fce9f3-3f05-4acf-a5a3-eac6ef4daa53) | Renvoie le taux d’escompte d’un titre |
| [Fonction DMAX](https://support.microsoft.com/office/f4e8209d-8958-4c3d-a1ee-6351665d41c2) | Renvoie la valeur maximale des entrées de base de données sélectionnées |
| [Fonction DMIN](https://support.microsoft.com/office/4ae6f1d9-1f26-40f1-a783-6dc3680192a3) | Renvoie la valeur minimale des entrées de base de données sélectionnées |
| [Fonctions DOLLAR, USDOLLAR](https://support.microsoft.com/office/a6cd05d9-9740-4ad3-a469-8109d18ff611) | Convertit un nombre en texte en utilisant le format monétaire $ (dollar) |
| [Fonction DOLLARDE](https://support.microsoft.com/office/db85aab0-1677-428a-9dfd-a38476693427) | Convertit un prix en dollars, exprimé sous forme de fraction, en un prix en dollars exprimé sous forme de nombre décimal |
| [Fonction DOLLARFR](https://support.microsoft.com/office/0835d163-3023-4a33-9824-3042c5d4f495) | Convertit un prix en dollars, exprimé sous forme de nombre décimal, en un prix en dollars exprimé sous forme de fraction |
| [Fonction DPRODUCT](https://support.microsoft.com/office/4f96b13e-d49c-47a7-b769-22f6d017cb31) | Multiplie les valeurs d’un champ particulier dans des enregistrements correspondant aux critères d’une base de données |
| [Fonction DSTDEV](https://support.microsoft.com/office/026b8c73-616d-4b5e-b072-241871c4ab96) | Calcule l’écart type en fonction d’un échantillon d’entrées de base de données sélectionnées |
| [Fonction DSTDEVP](https://support.microsoft.com/office/04b78995-da03-4813-bbd9-d74fd0f5d94b) | Calcule l’écart type en fonction de l’ensemble des entrées de base de données sélectionnées |
| [Fonction DSUM](https://support.microsoft.com/office/53181285-0c4b-4f5a-aaa3-529a322be41b) | Ajoute les nombres dans la colonne Champ des enregistrements de la base de données correspondant aux critères |
| [Fonction DURATION](https://support.microsoft.com/office/b254ea57-eadc-4602-a86a-c8e369334038) | Renvoie la durée annuelle d’un titre dont les intérêts sont perçus périodiquement |
| [Fonction Dlet](https://support.microsoft.com/office/d6747ca9-99c7-48bb-996e-9d7af00f3ed1) | Estime la variance en fonction d’un échantillon d’entrées de base de données sélectionnées |
| [Fonction DVARP](https://support.microsoft.com/office/eb0ba387-9cb7-45c8-81e9-0394912502fc) | Calcule la variance en fonction de l’ensemble des entrées de base de données sélectionnées |
| [Fonction EDATE](https://support.microsoft.com/office/3c920eb2-6e66-44e7-a1f5-753ae47ee4f5) | Renvoie le numéro de série de la date qui représente le nombre indiqué de mois précédant ou suivant la date de début |
| [Fonction EFFECT](https://support.microsoft.com/office/910d4e4c-79e2-4009-95e6-507e04f11bc4) | Renvoie le taux d’intérêt annuel effectif |
| [Fonction EOMONTH](https://support.microsoft.com/office/7314ffa1-2bc9-4005-9d66-f49db127d628) | Renvoie le numéro de série du dernier jour du mois précédant ou suivant un nombre de mois spécifié |
| [Fonction ERF](https://support.microsoft.com/office/c53c7e7b-5482-4b6c-883e-56df3c9af349) | Renvoie la valeur de la fonction d’erreur |
| [Fonction ERF.PRECISE](https://support.microsoft.com/office/9a349593-705c-4278-9a98-e4122831a8e0) | Renvoie la valeur de la fonction d’erreur |
| [Fonction ERFC](https://support.microsoft.com/office/736e0318-70ba-4e8b-8d08-461fe68b71b3) | Renvoie la valeur de la fonction d’erreur complémentaire |
| [Fonction ERFC.PRECISE](https://support.microsoft.com/office/e90e6bab-f45e-45df-b2ac-cd2eb4d4a273) | Renvoie la valeur de la fonction d’erreur complémentaire comprise entre x et l’infini |
| [Fonction ERROR.TYPE](https://support.microsoft.com/office/10958677-7c8d-44f7-ae77-b9a9ee6eefaa) | Renvoie un nombre correspondant à un type d’erreur |
| [Fonction EVEN](https://support.microsoft.com/office/197b5f06-c795-4c1e-8696-3c3b8a646cf9) | Arrondit un nombre au nombre entier pair supérieur |
| [Fonction EXACT](https://support.microsoft.com/office/d3087698-fc15-4a15-9631-12575cf29926) | Vérifie si deux valeurs textuelles sont identiques |
| [Fonction EXP](https://support.microsoft.com/office/c578f034-2c45-4c37-bc8c-329660a63abe) | Renvoie le nombre e élevé à la puissance d’un nombre donné |
| [Fonction EXPON.DIST](https://support.microsoft.com/office/4c12ae24-e563-4155-bf3e-8b78b6ae140e) | Renvoie la distribution exponentielle |
| [Fonction F.DIST](https://support.microsoft.com/office/a887efdc-7c8e-46cb-a74a-f884cd29b25d) | Renvoie la distribution de probabilité F |
| [Fonction F.DIST.RT](https://support.microsoft.com/office/d74cbb00-6017-4ac9-b7d7-6049badc0520) | Renvoie la distribution de probabilité F |
| [Fonction F.INV](https://support.microsoft.com/office/0dda0cf9-4ea0-42fd-8c3c-417a1ff30dbe) | Renvoie l’inverse de la distribution de probabilité F |
| [Fonction F.INVERT.RT](https://support.microsoft.com/office/d371aa8f-b0b1-40ef-9cc2-496f0693ac00) | Renvoie l’inverse de la distribution de probabilité F |
| [Fonction FACT](https://support.microsoft.com/office/ca8588c2-15f2-41c0-8e8c-c11bd471a4f3) | Renvoie la factorielle d’un nombre |
| [Fonction FACTDOUBLE](https://support.microsoft.com/office/e67697ac-d214-48eb-b7b7-cce2589ecac8) | Renvoie la factorielle double d’un nombre |
| [Fonction FALSE](https://support.microsoft.com/office/2d58dfa5-9c03-4259-bf8f-f0ae14346904) | Renvoie la valeur logique `FALSE` |
| [Fonctions FIND, FINDB](https://support.microsoft.com/office/c7912941-af2a-4bdf-a553-d0d89b0a0628) | Cherche une valeur textuelle dans une autre (en respectant la casse) |
| [Fonction FISHER](https://support.microsoft.com/office/d656523c-5076-4f95-b87b-7741bf236c69) | Renvoie la transformation de Fisher |
| [Fonction FISHERINV](https://support.microsoft.com/office/62504b39-415a-4284-a285-19c8e82f86bb) | Renvoie l’inverse de la transformation de Fisher |
| [Fonction FIXED](https://support.microsoft.com/office/ffd5723c-324c-45e9-8b96-e41be2a8274a) | Convertit un nombre en texte avec un nombre de décimales fixe |
| [Fonction FLOOR.MATH](https://support.microsoft.com/office/c302b599-fbdb-4177-ba19-2c2b1249a2f5) | Arrondit un nombre à l’entier ou au multiple inférieur le plus proche de l’argument de précision |
| [Fonction FLOOR.PRECISE](https://support.microsoft.com/office/f769b468-1452-4617-8dc3-02f842a0702e) | Arrondit un nombre à l’entier ou au multiple inférieur le plus proche de l’argument de précision. Quel que soit le signe du nombre, le nombre est arrondi à l’unité inférieure. |
| [Fonction FV](https://support.microsoft.com/office/2eef9f44-a084-4c61-bdd8-4fe4bb1b71b3) | Renvoie la valeur future d’un investissement |
| [Fonction FVSCHEDULE](https://support.microsoft.com/office/bec29522-bd87-4082-bab9-a241f3fb251d) | Renvoie la valeur future d’un investissement en appliquant une série de taux d’intérêt composites |
| [Fonction GAMMA](https://support.microsoft.com/office/ce1702b1-cf55-471d-8307-f83be0fc5297) | Renvoie la valeur de la fonction Gamma |
| [Fonction GAMMA.DIST](https://support.microsoft.com/office/9b6f1538-d11c-4d5f-8966-21f6a2201def) | Renvoie la distribution suivant une loi Gamma |
| [Fonction GAMMA.INV](https://support.microsoft.com/office/74991443-c2b0-4be5-aaab-1aa4d71fbb18) | Renvoie l’inverse de la distribution cumulée suivant une loi Gamma |
| [Fonction GAMMALN](https://support.microsoft.com/office/b838c48b-c65f-484f-9e1d-141c55470eb9) | Renvoie le logarithme népérien de la fonction gamma, Γ(x) |
| [Fonction GAMMALN.PRECISE](https://support.microsoft.com/office/5cdfe601-4e1e-4189-9d74-241ef1caa599) | Renvoie le logarithme népérien de la fonction gamma, Γ(x) |
| [Fonction GAUSS](https://support.microsoft.com/office/069f1b4e-7dee-4d6a-a71f-4b69044a6b33) | Renvoie 0,5 de moins que la distribution cumulée suivant une loi normale centrée réduite |
| [Fonction GCD](https://support.microsoft.com/office/d5107a51-69e3-461f-8e4c-ddfc21b5073a) | Renvoie le plus grand diviseur commun |
| [Fonction GEOMEAN](https://support.microsoft.com/office/db1ac48d-25a5-40a0-ab83-0b38980e40d5) | Renvoie la moyenne géométrique |
| [Fonction GESTEP](https://support.microsoft.com/office/f37e7d2a-41da-4129-be95-640883fca9df) | Vérifie si un nombre est supérieur à une valeur seuil |
| [Fonction HARMEAN](https://support.microsoft.com/office/5efd9184-fab5-42f9-b1d3-57883a1d3bc6) | Renvoie la moyenne harmonique |
| [Fonction HEX2BIN](https://support.microsoft.com/office/a13aafaa-5737-4920-8424-643e581828c1) | Convertit un nombre hexadécimal en nombre binaire |
| [Fonction HEX2DEC](https://support.microsoft.com/office/8c8c3155-9f37-45a5-a3ee-ee5379ef106e) | Convertit un nombre hexadécimal en nombre décimal |
| [Fonction HEX2OCT](https://support.microsoft.com/office/54d52808-5d19-4bd0-8a63-1096a5d11912) | Convertit un nombre hexadécimal en nombre octal |
| [Fonction HLOOKUP](https://support.microsoft.com/office/a3034eec-b719-4ba3-bb65-e1ad662ed95f) | Cherche dans la première ligne d’un tableau et renvoie la valeur de la cellule indiquée |
| [Fonction HOUR](https://support.microsoft.com/office/a3afa879-86cb-4339-b1b5-2dd2d7310ac7) | Convertit un numéro de série en heure |
| [Fonction HYPERLINK](https://support.microsoft.com/office/333c7ce6-c5ae-4164-9c47-7de9b76f577f) | Crée un raccourci ou un renvoi qui ouvre un document stocké sur un serveur réseau, un intranet ou Internet |
| [Fonction HYPGEOM.DIST](https://support.microsoft.com/office/6dbd547f-1d12-4b1f-8ae5-b0d9e3d22fbf) | Renvoie la distribution suivant une loi hypergéométrique |
| [Fonction IF](https://support.microsoft.com/office/69aed7c9-4e8a-4755-a9bc-aa8bbff73be2) | Indique un test logique à effectuer |
| [Fonction IMABS](https://support.microsoft.com/office/b31e73c6-d90c-4062-90bc-8eb351d765a1) | Renvoie la valeur absolue (module) d’un nombre complexe |
| [Fonction IMAGINARY](https://support.microsoft.com/office/dd5952fd-473d-44d9-95a1-9a17b23e428a) | Renvoie le coefficient imaginaire d’un nombre complexe |
| [Fonction IMARGUMENT](https://support.microsoft.com/office/eed37ec1-23b3-4f59-b9f3-d340358a034a) | Renvoie l’argument thêta, un angle exprimé en radians |
| [Fonction IMCONJUGATE](https://support.microsoft.com/office/2e2fc1ea-f32b-4f9b-9de6-233853bafd42) | Renvoie le conjugué complexe d’un nombre complexe |
| [Fonction IMCOS](https://support.microsoft.com/office/dad75277-f592-4a6b-ad6c-be93a808a53c) | Renvoie le cosinus d’un nombre complexe |
| [Fonction IMCOSH](https://support.microsoft.com/office/053e4ddb-4122-458b-be9a-457c405e90ff) | Renvoie le cosinus hyperbolique d’un nombre complexe |
| [Fonction IMCOT](https://support.microsoft.com/office/dc6a3607-d26a-4d06-8b41-8931da36442c) | Renvoie la cotangente d’un nombre complexe |
| [Fonction IMCSC](https://support.microsoft.com/office/9e158d8f-2ddf-46cd-9b1d-98e29904a323) | Renvoie la cosécante d’un nombre complexe |
| [Fonction IMCSCH](https://support.microsoft.com/office/c0ae4f54-5f09-4fef-8da0-dc33ea2c5ca9) | Renvoie la cosécante hyperbolique d’un nombre complexe |
| [Fonction IMDIV](https://support.microsoft.com/office/a505aff7-af8a-4451-8142-77ec3d74d83f) | Renvoie le quotient de deux nombres complexes |
| [Fonction IMEXP](https://support.microsoft.com/office/c6f8da1f-e024-4c0c-b802-a60e7147a95f) | Renvoie la fonction exponentielle d’un nombre complexe |
| [Fonction IMLN](https://support.microsoft.com/office/32b98bcf-8b81-437c-a636-6fb3aad509d8) | Renvoie le logarithme népérien d’un nombre complexe |
| [Fonction IMLOG10](https://support.microsoft.com/office/58200fca-e2a2-4271-8a98-ccd4360213a5) | Calcule le logarithme d’un nombre complexe en base 10 |
| [Fonction IMLOG2](https://support.microsoft.com/office/152e13b4-bc79-486c-a243-e6a676878c51) | Calcule le logarithme d’un nombre complexe en base 2 |
| [Fonction IMPOWER](https://support.microsoft.com/office/210fd2f5-f8ff-4c6a-9d60-30e34fbdef39) | Renvoie un nombre complexe élevé à une puissance entière |
| [Fonction IMPRODUCT](https://support.microsoft.com/office/2fb8651a-a4f2-444f-975e-8ba7aab3a5ba) | Renvoie le produit de 2 à 255 nombres complexes |
| [Fonction IMREAL](https://support.microsoft.com/office/d12bc4c0-25d0-4bb3-a25f-ece1938bf366) | Renvoie le coefficient réel d’un nombre complexe |
| [Fonction IMSEC](https://support.microsoft.com/office/6df11132-4411-4df4-a3dc-1f17372459e0) | Renvoie la sécante d’un nombre complexe |
| [Fonction IMSECH](https://support.microsoft.com/office/f250304f-788b-4505-954e-eb01fa50903b) | Renvoie la sécante hyperbolique d’un nombre complexe |
| [Fonction IMSIN](https://support.microsoft.com/office/1ab02a39-a721-48de-82ef-f52bf37859f6) | Renvoie le sinus d’un nombre complexe |
| [Fonction IMSINH](https://support.microsoft.com/office/dfb9ec9e-8783-4985-8c42-b028e9e8da3d) | Renvoie le sinus hyperbolique d’un nombre complexe |
| [Fonction IMSQRT](https://support.microsoft.com/office/e1753f80-ba11-4664-a10e-e17368396b70) | Renvoie la racine carrée d’un nombre complexe |
| [Fonction IMSUB](https://support.microsoft.com/office/2e404b4d-4935-4e85-9f52-cb08b9a45054) | Renvoie la différence entre deux nombres complexes |
| [Fonction IMSUM](https://support.microsoft.com/office/81542999-5f1c-4da6-9ffe-f1d7aaa9457f) | Renvoie la somme de plusieurs nombres complexes |
| [Fonction IMTAN](https://support.microsoft.com/office/8478f45d-610a-43cf-8544-9fc0b553a132) | Renvoie la tangente d’un nombre complexe |
| [Fonction INT](https://support.microsoft.com/office/a6c4af9e-356d-4369-ab6a-cb1fd9d343ef) | Arrondit un nombre à l’entier inférieur le plus proche |
| [Fonction INTRATE](https://support.microsoft.com/office/5cb34dde-a221-4cb6-b3eb-0b9e55e1316f) | Renvoie le taux d’intérêt pour un titre totalement investi |
| [Fonction IPMT](https://support.microsoft.com/office/5cce0ad6-8402-4a41-8d29-61a0b054cb6f) | Renvoie le montant des intérêts d’un investissement pour une période donnée |
| [Fonction IRR](https://support.microsoft.com/office/64925eaa-9988-495b-b290-3ad0c163c1bc) | Renvoie le taux de rendement interne pour une série de mouvements de trésorerie |
| [Fonction ISERR](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Renvoie `TRUE` si la valeur est une valeur d’erreur, sauf #N/A |
| [Fonction ISERROR](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Renvoie `TRUE` si la valeur est une valeur d’erreur |
| [Fonction ISEVEN](https://support.microsoft.com/office/aa15929a-d77b-4fbb-92f4-2f479af55356) | Renvoie `TRUE` si le nombre est pair |
| [Fonction ISFORMULA](https://support.microsoft.com/office/e4d1355f-7121-4ef2-801e-3839bfd6b1e5) | Renvoie `TRUE` s’il existe une référence à une cellule qui contient une formule |
| [Fonction ISLOGICAL](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Renvoie `TRUE` si la valeur est une valeur logique |
| [Fonction ISNA](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Renvoie `TRUE` si la valeur est la valeur d’erreur #N/A |
| [Fonction ISNONTEXT](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Renvoie `TRUE` si la valeur n’est pas textuelle |
| [Fonction ISNUMBER](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Renvoie `TRUE` si la valeur est un nombre |
| [Fonction ISO.CEILING](https://support.microsoft.com/office/e587bb73-6cc2-4113-b664-ff5b09859a83) | Renvoie un nombre arrondi à l’entier ou au multiple supérieur le plus proche de l’argument de précision |
| [Fonction ISODD](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Renvoie `TRUE` si le nombre est impair |
| [Fonction ISOWEEKNUM](https://support.microsoft.com/office/1c2d0afe-d25b-4ab1-8894-8d0520e90e0e) | Renvoie le numéro de la semaine ISO de l’année pour une date donnée |
| [Fonction ISPMT](https://support.microsoft.com/office/fa58adb6-9d39-4ce0-8f43-75399cea56cc) | Calcule le montant des intérêts payés au cours d’une période spécifique d’un investissement |
| [Fonction ISREF](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Renvoie `TRUE` si la valeur est une référence |
| [Fonction ISTEXT](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Renvoie `TRUE` si la valeur est textuelle |
| [Fonction KURT](https://support.microsoft.com/office/bc3a265c-5da4-4dcb-b7fd-c237789095ab) | Renvoie le kurtosis d’un jeu de données |
| [Fonction LARGE](https://support.microsoft.com/office/3af0af19-1190-42bb-bb8b-01672ec00a64) | Renvoie la k-ième plus grande valeur d’un jeu de données |
| [Fonction LCM](https://support.microsoft.com/office/7152b67a-8bb5-4075-ae5c-06ede5563c94) | Renvoie le plus petit dénominateur commun |
| [Fonctions LEFT, LEFTB](https://support.microsoft.com/office/9203d2d2-7960-479b-84c6-1ea52b99640c) | Renvoie les caractères les plus à gauche d’une valeur textuelle |
| [Fonctions LEN, LENB](https://support.microsoft.com/office/29236f94-cedc-429d-affd-b5e33d2c67cb) | Renvoie le nombre de caractères dans une chaîne de texte |
| [Fonction LN](https://support.microsoft.com/office/81fe1ed7-dac9-4acd-ba1d-07a142c6118f) | Renvoie le logarithme népérien d’un nombre |
| [Fonction LOG](https://support.microsoft.com/office/4e82f196-1ca9-4747-8fb0-6c4a3abb3280) | Renvoie le logarithme d’un nombre selon la base spécifiée |
| [Fonction LOG10](https://support.microsoft.com/office/c75b881b-49dd-44fb-b6f4-37e3486a0211) | Renvoie le logarithme d’un nombre en base 10 |
| [Fonction LOGNORM.DIST](https://support.microsoft.com/office/eb60d00b-48a9-4217-be2b-6074aee6b070) | Renvoie la distribution suivant une loi lognormale cumulée |
| [Fonction LOGNORM.INV](https://support.microsoft.com/office/fe79751a-f1f2-4af8-a0a1-e151b2d4f600) | Renvoie l’inverse de la distribution cumulée suivant une loi lognormale |
| [Fonction LOOKUP](https://support.microsoft.com/office/446d94af-663b-451d-8251-369d5e3864cb) | Cherche des valeurs dans un vecteur ou un tableau |
| [Fonction LOWER](https://support.microsoft.com/office/3f21df02-a80c-44b2-afaf-81358f9fdeb4) | Convertit le texte en minuscules |
| [Fonction MATCH](https://support.microsoft.com/office/e8dffd45-c762-47d6-bf89-533f4a37673a) | Cherche des valeurs dans une référence ou un tableau |
| [Fonction MAX](https://support.microsoft.com/office/e0012414-9ac8-4b34-9a47-73e662c08098) | Renvoie la valeur maximale contenue dans une liste d’arguments |
| [Fonction MAXA](https://support.microsoft.com/office/814bda1e-3840-4bff-9365-2f59ac2ee62d) | Renvoie la valeur maximale contenue dans une liste d’arguments, y compris les nombres, le texte et les valeurs logiques |
| [Fonction MDURATION](https://support.microsoft.com/office/b3786a69-4f20-469a-94ad-33e5b90a763c) | Renvoie la durée modifiée de Macauley pour un titre avec une valeur estimée à 100 dollars |
| [Fonction MEDIAN](https://support.microsoft.com/office/d0916313-4753-414c-8537-ce85bdd967d2) | Renvoie la valeur médiane des nombres donnés |
| [Fonction MID, MIDB](https://support.microsoft.com/office/d5f9e25c-d7d6-472e-b568-4ecb12433028) | Renvoie un nombre déterminé de caractères d’une chaîne de texte en commençant à la position indiquée |
| [Fonction MIN](https://support.microsoft.com/office/61635d12-920f-4ce2-a70f-96f202dcc152) | Renvoie la valeur minimale contenue dans une liste d’arguments |
| [Fonction MINA](https://support.microsoft.com/office/245a6f46-7ca5-4dc7-ab49-805341bc31d3) | Renvoie la plus petite valeur contenue dans une liste d’arguments, y compris les nombres, le texte et les valeurs logiques |
| [Fonction MINUTE](https://support.microsoft.com/office/af728df0-05c4-4b07-9eed-a84801a60589) | Convertit un numéro de série en minute |
| [Fonction MIRR](https://support.microsoft.com/office/b020f038-7492-4fb4-93c1-35c345b53524) | Renvoie le taux de rendement interne lorsque des mouvements de trésorerie positifs et négatifs sont financés à des taux différents |
| [Fonction MOD](https://support.microsoft.com/office/9b6cd169-b6ee-406a-a97b-edf2a9dc24f3) | Renvoie le reste d’une division |
| [Fonction MONTH](https://support.microsoft.com/office/579a2881-199b-48b2-ab90-ddba0eba86e8) | Convertit un numéro de série en mois |
| [Fonction MROUND](https://support.microsoft.com/office/c299c3b0-15a5-426d-aa4b-d2d5b3baf427) | Renvoie un nombre arrondi au dénominateur souhaité |
| [Fonction MULTINOMIAL](https://support.microsoft.com/office/6fa6373c-6533-41a2-a45e-a56db1db1bf6) | Calcule la multinomiale d’un ensemble de nombres |
| [Fonction N](https://support.microsoft.com/office/a624cad1-3635-4208-b54a-29733d1278c9) | Renvoie une valeur convertie en nombre |
| [Fonction NA](https://support.microsoft.com/office/5469c2d1-a90c-4fb5-9bbc-64bd9bb6b47c) | Renvoie la valeur d’erreur #N/A |
| [Fonction NEGBINOM.DIST](https://support.microsoft.com/office/c8239f89-c2d0-45bd-b6af-172e570f8599) | Renvoie la distribution négative binomiale |
| [Fonction NETWORKDAYS](https://support.microsoft.com/office/48e717bf-a7a3-495f-969e-5005e3eb18e7) | Renvoie le nombre de jours ouvrés entiers entre deux dates |
| [Fonction NETWORKDAYS.INTL](https://support.microsoft.com/office/a9b26239-4f20-46a1-9ab8-4e925bfd5e28) | Renvoie le nombre de jours ouvrés entiers compris entre deux dates à l’aide de paramètres indiquant le nombre de jours compris dans un week-end |
| [Fonction NOMINAL](https://support.microsoft.com/office/7f1ae29b-6b92-435e-b950-ad8b190ddd2b) | Renvoie le taux d’intérêt nominal annuel |
| [Fonction NORM.DIST](https://support.microsoft.com/office/edb1cc14-a21c-4e53-839d-8082074c9f8d) | Renvoie la distribution cumulée suivant une loi normale |
| [Fonction NORM.INV](https://support.microsoft.com/office/54b30935-fee7-493c-bedb-2278a9db7e13) | Renvoie l’inverse de la distribution cumulée suivant une loi normale |
| [Fonction NORM.S.DIST](https://support.microsoft.com/office/1e787282-3832-4520-a9ae-bd2a8d99ba88) | Renvoie la distribution cumulée suivant une loi normale centrée réduite |
| [Fonction NORM.S.INV](https://support.microsoft.com/office/d6d556b4-ab7f-49cd-b526-5a20918452b1) | Renvoie l’inverse de la distribution cumulée suivant une loi normale centrée réduite |
| [Fonction NOT](https://support.microsoft.com/office/9cfc6011-a054-40c7-a140-cd4ba2d87d77) | Inverse la logique de son argument |
| [Fonction NOW](https://support.microsoft.com/office/3337fd29-145a-4347-b2e6-20c904739c46) | Renvoie le numéro de série de la date et de l’heure actuelles |
| [Fonction NPER](https://support.microsoft.com/office/240535b5-6653-4d2d-bfcf-b6a38151d815) | Renvoie le nombre de paiements d’un investissement |
| [Fonction NPV](https://support.microsoft.com/office/8672cb67-2576-4d07-b67b-ac28acf2a568) | Renvoie la valeur nette actuelle d’un investissement, en fonction d’une série de flux de trésorerie périodiques et d’un taux d’escompte |
| [Fonction NUMBERVALUE](https://support.microsoft.com/office/1b05c8cf-2bfa-4437-af70-596c7ea7d879) | Convertit le texte en nombre quels que soient les paramètres régionaux |
| [Fonction OCT2BIN](https://support.microsoft.com/office/55383471-3c56-4d27-9522-1a8ec646c589) | Convertit un nombre octal en nombre binaire |
| [Fonction OCT2DEC](https://support.microsoft.com/office/87606014-cb98-44b2-8dbb-e48f8ced1554) | Convertit un nombre octal en nombre décimal |
| [Fonction OCT2HEX](https://support.microsoft.com/office/912175b4-d497-41b4-a029-221f051b858f) | Convertit un nombre octal en nombre hexadécimal |
| [Fonction ODD](https://support.microsoft.com/office/deae64eb-e08a-4c88-8b40-6d0b42575c98) | Arrondit un nombre à l’entier impair supérieur le plus proche |
| [Fonction ODDFPRICE](https://support.microsoft.com/office/d7d664a8-34df-4233-8d2b-922bcf6a69e1) | Renvoie le prix par valeur faciale de 100 dollars d’un titre dont la première période est irrégulière |
| [Fonction ODDFYIELD](https://support.microsoft.com/office/66bc8b7b-6501-4c93-9ce3-2fd16220fe37) | Renvoie le rendement d’un titre dont la première période est irrégulière |
| [Fonction ODDLPRICE](https://support.microsoft.com/office/fb657749-d200-4902-afaf-ed5445027fc4) | Renvoie le prix par valeur faciale de 100 dollars d’un titre dont la dernière période est irrégulière |
| [Fonction ODDLYIELD](https://support.microsoft.com/office/c873d088-cf40-435f-8d41-c8232fee9238) | Renvoie le rendement d’un titre dont la dernière période est irrégulière |
| [Fonction OR](https://support.microsoft.com/office/7d17ad14-8700-4281-b308-00b131e22af0) | Renvoie `TRUE` si un argument a la valeur True |
| [Fonction PDURATION](https://support.microsoft.com/office/44f33460-5be5-4c90-b857-22308892adaf) | Renvoie le nombre de périodes requises par un investissement pour atteindre une valeur spécifiée |
| [Fonction PERCENTILE.EXC](https://support.microsoft.com/office/bbaa7204-e9e1-4010-85bf-c31dc5dce4ba) | Renvoie le k-ième centile de valeur d’une plage, où k se trouve dans la plage de 0 à 1 exclus |
| [Fonction PERCENTILE.INC](https://support.microsoft.com/office/680f9539-45eb-410b-9a5e-c1355e5fe2ed) | Renvoie le k-ième centile des valeurs d’une plage |
| [Fonction PERCENTRANK.EXC](https://support.microsoft.com/office/d8afee96-b7e2-4a2f-8c01-8fcdedaa6314) | Renvoie le rang d’une valeur dans un ensemble de données défini comme pourcentage (0..1, exclus) de cet ensemble |
| [Fonction PERCENTRANK.INC](https://support.microsoft.com/office/149592c9-00c0-49ba-86c1-c1f45b80463a) | Renvoie le rang en pourcentage d’une valeur dans un jeu de données |
| [Fonction PERMUT](https://support.microsoft.com/office/3bd1cb9a-2880-41ab-a197-f246a7a602d3) | Renvoie le nombre de permutations pour un nombre d’objets donné |
| [Fonction PERMUTATIONA](https://support.microsoft.com/office/6c7d7fdc-d657-44e6-aa19-2857b25cae4e) | Renvoie le nombre de permutations pour un nombre d’objets donné (avec répétitions) pouvant être sélectionnés à partir du nombre total d’objets |
| [Fonction PHI](https://support.microsoft.com/office/23e49bc6-a8e8-402d-98d3-9ded87f6295c) | Renvoie la valeur de la fonction de densité pour une distribution suivant une loi normale centrée réduite |
| [Fonction PI](https://support.microsoft.com/office/264199d0-a3ba-46b8-975a-c4a04608989b) | Renvoie la valeur de pi |
| [Fonction PMT](https://support.microsoft.com/office/0214da64-9a63-4996-bc20-214433fa6441) | Renvoie le montant périodique d’une annuité |
| [Fonction POISSON.DIST](https://support.microsoft.com/office/8fe148ff-39a2-46cb-abf3-7772695d9636) | Renvoie la distribution suivant une loi de Poisson |
| [Fonction POWER](https://support.microsoft.com/office/d3f2908b-56f4-4c3f-895a-07fb519c362a) | Renvoie le résultat d’un nombre élevé à une puissance |
| [Fonction PPMT](https://support.microsoft.com/office/c370d9e3-7749-4ca4-beea-b06c6ac95e1b) | Renvoie la part de remboursement du principal d’un emprunt pour une période donnée |
| [Fonction PRICE](https://support.microsoft.com/office/3ea9deac-8dfa-436f-a7c8-17ea02c21b0a) | Renvoie le prix par valeur faciale de 100 dollars d’un titre dont les intérêts sont payés périodiquement |
| [Fonction PRICEDISC](https://support.microsoft.com/office/d06ad7c1-380e-4be7-9fd9-75e3079acfd3) | Renvoie le prix par valeur faciale de 100 dollars pour un titre escompté |
| [Fonction PRICEMAT](https://support.microsoft.com/office/52c3b4da-bc7e-476a-989f-a95f675cae77) | Renvoie le prix par valeur faciale de 100 dollars d’un titre dont les intérêts sont payés à échéance |
| [PRODUIT, fonction](https://support.microsoft.com/office/8e6b5b24-90ee-4650-aeec-80982a0512ce) | Multiplie ses arguments |
| [Fonction PROPER](https://support.microsoft.com/office/52a5a283-e8b2-49be-8506-b2887b889f94) | Met en majuscule la première lettre de chaque mot d’une valeur textuelle |
| [Fonction PV](https://support.microsoft.com/office/23879d31-0e02-4321-be01-da16e8168cbd) | Renvoie la valeur actuelle d’un investissement |
| [Fonction QUARTILE.EXC](https://support.microsoft.com/office/5a355b7a-840b-4a01-b0f1-f538c2864cad) | Renvoie le quartile de l’ensemble de données d’après des valeurs de centile comprises entre 0 et 1, exclus |
| [Fonction QUARTILE.INC](https://support.microsoft.com/office/1bbacc80-5075-42f1-aed6-47d735c4819d) | Renvoie le quartile d’un jeu de données |
| [Fonction QUOTIENT](https://support.microsoft.com/office/9f7bf099-2a18-4282-8fa4-65290cc99dee) | Renvoie la partie entière d’une division |
| [Fonction RADIANS](https://support.microsoft.com/office/ac409508-3d48-45f5-ac02-1497c92de5bf) | Convertit des degrés en radians |
| [Fonction RAND](https://support.microsoft.com/office/4cbfa695-8869-4788-8d90-021ea9f5be73) | Renvoie un nombre aléatoire compris entre 0 et 1 |
| [Fonction RANDBETWEEN](https://support.microsoft.com/office/4cc7f0d1-87dc-4eb7-987f-a469ab381685) | Renvoie un nombre aléatoire entre les nombres que vous spécifiez |
| [Fonction RANK.AVG](https://support.microsoft.com/office/bd406a6f-eb38-4d73-aa8e-6d1c3c72e83a) | Renvoie le rang d’un nombre dans une liste de nombres |
| [Fonction RANK.EQ](https://support.microsoft.com/office/284858ce-8ef6-450e-b662-26245be04a40) | Renvoie le rang d’un nombre dans une liste de nombres |
| [Fonction RATE](https://support.microsoft.com/office/9f665657-4a7e-4bb7-a030-83fc59e748ce) | Renvoie le taux d’intérêt par période pour une annuité |
| [Fonction RECEIVED](https://support.microsoft.com/office/7a3f8b93-6611-4f81-8576-828312c9b5e5) | Renvoie le montant reçu lorsqu’un titre totalement investi arrive à échéance |
| [Fonctions REPLACE, REPLACEB](https://support.microsoft.com/office/8d799074-2425-4a8a-84bc-82472868878a) | Remplace des caractères dans un texte |
| [Fonction REPT](https://support.microsoft.com/office/04c4d778-e712-43b4-9c15-d656582bb061) | Répète un texte un certain nombre de fois |
| [Fonctions RIGHT, RIGHTB](https://support.microsoft.com/office/240267ee-9afa-4639-a02b-f19e1786cf2f) | Renvoie les caractères les plus à droite d’une valeur textuelle |
| [Fonction ROMAN](https://support.microsoft.com/office/d6b0b99e-de46-4704-a518-b45a0f8b56f5) | Convertit un numéral arabe en caractères romans, sous forme de texte |
| [Fonction ROUND](https://support.microsoft.com/office/c018c5d8-40fb-4053-90b1-b3e7f61a213c) | Arrondit un nombre à un nombre de chiffres spécifié |
| [Fonction ROUNDDOWN](https://support.microsoft.com/office/2ec94c73-241f-4b01-8c6f-17e6d7968f53) | Arrondit un nombre à la valeur d’arrondi la plus proche de zéro |
| [Fonction ROUNDUP](https://support.microsoft.com/office/f8bc9b23-e795-47db-8703-db171d0c42a7) | Arrondit un nombre à la valeur d’arrondi la plus éloignée de zéro |
| [Fonction ROWS](https://support.microsoft.com/office/b592593e-3fc2-47f2-bec1-bda493811597) | Renvoie le nombre de lignes dans une référence |
| [Fonction RRI](https://support.microsoft.com/office/6f5822d8-7ef1-4233-944c-79e8172930f4) | Renvoie un taux d’intérêt équivalent pour la croissance d’un investissement |
| [Fonction SEC](https://support.microsoft.com/office/ff224717-9c87-4170-9b58-d069ced6d5f7) | Renvoie la sécante d’un angle |
| [Fonction SECH](https://support.microsoft.com/office/e05a789f-5ff7-4d7f-984a-5edb9b09556f) | Renvoie la sécante hyperbolique d’un angle |
| [Fonction SECOND](https://support.microsoft.com/office/740d1cfc-553c-4099-b668-80eaa24e8af1) | Convertit un numéro de série en seconde |
| [Fonction SERIESSUM](https://support.microsoft.com/office/a3ab25b5-1093-4f5b-b084-96c49087f637) | Renvoie le total d’une série de puissance basé sur la formule |
| [Fonction SHEET](https://support.microsoft.com/office/44718b6f-8b87-47a1-a9d6-b701c06cff24) | Renvoie le numéro de la feuille référencée |
| [Fonction SHEETS](https://support.microsoft.com/office/770515eb-e1e8-45ce-8066-b557e5e4b80b) | Renvoie le nombre de feuilles dans une référence |
| [Fonction SIGN](https://support.microsoft.com/office/109c932d-fcdc-4023-91f1-2dd0e916a1d8) | Renvoie le signe d’un nombre |
| [Fonction SIN](https://support.microsoft.com/office/cf0e3432-8b9e-483c-bc55-a76651c95602) | Renvoie le sinus d’un angle donné |
| [Fonction SINH](https://support.microsoft.com/office/1e4e8b9f-2b65-43fc-ab8a-0a37f4081fa7) | Renvoie le sinus hyperbolique d’un nombre |
| [Fonction SKEW](https://support.microsoft.com/office/bdf49d86-b1ef-4804-a046-28eaea69c9fa) | Renvoie l’asymétrie d’une distribution |
| [Fonction SKEW.P](https://support.microsoft.com/office/76530a5c-99b9-48a1-8392-26632d542fcb) | Renvoie l’asymétrie d’une distribution en fonction d’une population : la caractérisation du degré d’asymétrie d’une distribution par rapport à sa moyenne |
| [Fonction SLN](https://support.microsoft.com/office/cdb666e5-c1c6-40a7-806a-e695edc2f1c8) | Renvoie l’amortissement linéaire d’une immobilisation pour une période |
| [Fonction SMALL](https://support.microsoft.com/office/17da8222-7c82-42b2-961b-14c45384df07) | Renvoie la k-ième plus petite valeur d’un jeu de données |
| [Fonction SQRT](https://support.microsoft.com/office/654975c2-05c4-4831-9a24-2c65e4040fdf) | Renvoie une racine carrée positive |
| [Fonction SQRTPI](https://support.microsoft.com/office/1fb4e63f-9b51-46d6-ad68-b3e7a8b519b4) | Renvoie la racine carrée de (nombre * pi). |
| [Fonction STANDARDIZE](https://support.microsoft.com/office/81d66554-2d54-40ec-ba83-6437108ee775) | Renvoie une valeur normalisée |
| [Fonction STDEV.P](https://support.microsoft.com/office/6e917c05-31a0-496f-ade7-4f4e7462f285) | Calcule l’écart type en fonction de la population entière |
| [Fonction STDEV.S](https://support.microsoft.com/office/7d69cf97-0c1f-4acf-be27-f3e83904cc23) | Évalue l’écart type en fonction d’un échantillon |
| [Fonction STDEVA](https://support.microsoft.com/office/5ff38888-7ea5-48de-9a6d-11ed73b29e9d) | Évalue l’écart type en fonction d’un échantillon, y compris les nombres, le texte et les valeurs logiques |
| [Fonction STDEVPA](https://support.microsoft.com/office/5578d4d6-455a-4308-9991-d405afe2c28c) | Calcule l’écart type en fonction de la population entière, y compris les nombres, le texte et les valeurs logiques |
| [Fonction SUBSTITUTE](https://support.microsoft.com/office/6434944e-a904-4336-a9b0-1e58df3bc332) | Remplace le nouveau texte par l’ancien texte d’une chaîne de texte |
| [Fonction SUBTOTAL](https://support.microsoft.com/office/7b027003-f060-4ade-9040-e478765b9939) | Renvoie un sous-total dans une liste ou une base de données |
| [Fonction SUM](https://support.microsoft.com/office/043e1c7d-7726-4e80-8f32-07b23e057f89) | Ajoute ses arguments |
| [Fonction SUMIF](https://support.microsoft.com/office/169b8c99-c05c-4483-a712-1697a653039b) | Ajoute les cellules spécifiées par un critère donné |
| [Fonction SUMIFS](https://support.microsoft.com/office/c9e748f5-7ea7-455d-9406-611cebce642b) | Ajoute les cellules d’une plage répondant à plusieurs critères |
| [Fonction SUMSQ](https://support.microsoft.com/office/e3313c02-51cc-4963-aae6-31442d9ec307) | Renvoie le total des carrés des arguments |
| [Fonction SYD](https://support.microsoft.com/office/069f8106-b60b-4ca2-98e0-2a0f206bdb27) | Renvoie l’amortissement des chiffres cumulés sur l’année d’une immobilisation pour une période spécifique |
| [Fonction T](https://support.microsoft.com/office/fb83aeec-45e7-4924-af95-53e073541228) | Convertit ses arguments en texte |
| [Fonction T.DIST](https://support.microsoft.com/office/4329459f-ae91-48c2-bba8-1ead1c6c21b2) | Renvoie les points de pourcentage (probabilité) pour la distribution suivant la loi T de Student |
| [Fonction T.DIST.2T](https://support.microsoft.com/office/198e9340-e360-4230-bd21-f52f22ff5c28) | Renvoie les points de pourcentage (probabilité) pour la distribution suivant la loi T de Student |
| [Fonction T.DIST.RT](https://support.microsoft.com/office/20a30020-86f9-4b35-af1f-7ef6ae683eda) | Renvoie la distribution suivant la loi T de Student |
| [Fonction T.INV](https://support.microsoft.com/office/2908272b-4e61-4942-9df9-a25fec9b0e2e) | Renvoie la valeur t de la distribution suivant la loi T de Student sous forme de fonction de probabilité et de degrés de liberté |
| [Fonction T.INV.2T](https://support.microsoft.com/office/ce72ea19-ec6c-4be7-bed2-b9baf2264f17) | Renvoie l’inverse de la distribution suivant la loi T de Student |
| [Fonction TAN](https://support.microsoft.com/office/08851a40-179f-4052-b789-d7f699447401) | Renvoie la tangente d’un nombre |
| [Fonction TANH](https://support.microsoft.com/office/017222f0-a0c3-4f69-9787-b3202295dc6c) | Renvoie la tangente hyperbolique d’un nombre |
| [Fonction TBILLEQ](https://support.microsoft.com/office/2ab72d90-9b4d-4efe-9fc2-0f81f2c19c8c) | Renvoie le rapport lié aux titres pour un bon du Trésor |
| [Fonction TBILLPRICE](https://support.microsoft.com/office/eacca992-c29d-425a-9eb8-0513fe6035a2) | Renvoie le prix par valeur faciale de 100 dollars pour un bon du Trésor |
| [Fonction TBILLYIELD](https://support.microsoft.com/office/6d381232-f4b0-4cd5-8e97-45b9c03468ba) | Renvoie le rapport pour un bon du Trésor |
| [Fonction TEXT](https://support.microsoft.com/office/20d5ac4d-7b94-49fd-bb38-93d29371225c) | Met en forme un nombre et le convertit en texte |
| [Fonction TIME](https://support.microsoft.com/office/9a5aff99-8f7d-4611-845e-747d0b8d5457) | Renvoie le numéro de série d’une heure précise |
| [Fonction TIMEVALUE](https://support.microsoft.com/office/0b615c12-33d8-4431-bf3d-f3eb6d186645) | Convertit une heure au format texte en numéro de série |
| [Fonction TODAY](https://support.microsoft.com/office/5eb3078d-a82c-4736-8930-2f51a028fdd9) | Renvoie le numéro de série de la date du jour |
| [Fonction TRIM](https://support.microsoft.com/office/410388fa-c5df-49c6-b16c-9e5630b479f9) | Supprime les espaces du texte |
| [Fonction TRIMMEAN](https://support.microsoft.com/office/d90c9878-a119-4746-88fa-63d988f511d3) | Renvoie la moyenne de la partie intérieure d’un jeu de données |
| [Fonction TRUE](https://support.microsoft.com/office/7652c6e3-8987-48d0-97cd-ef223246b3fb) | Renvoie la valeur logique `TRUE` |
| [Fonction TRUNC](https://support.microsoft.com/office/8b86a64c-3127-43db-ba14-aa5ceb292721) | Tronque un nombre en entier |
| [Fonction TYPE](https://support.microsoft.com/office/45b4e688-4bc3-48b3-a105-ffa892995899) | Renvoie un nombre indiquant le type de données d’une valeur |
| [Fonction UNICHAR](https://support.microsoft.com/office/ffeb64f5-f131-44c6-b332-5cd72f0659b8) | Renvoie le caractère unicode référencé par la valeur numérique donnée |
| [Fonction UNICODE](https://support.microsoft.com/office/adb74aaa-a2a5-4dde-aff6-966e4e81f16f) | Renvoie le nombre (point de code) qui correspond au premier caractère du texte |
| [Fonction UPPER](https://support.microsoft.com/office/c11f29b3-d1a3-4537-8df6-04d0049963d6) | Convertit le texte en majuscules |
| [Fonction VALUE](https://support.microsoft.com/office/257d0108-07dc-437d-ae1c-bc2d3953d8c2) | Convertit un argument textuel en nombre |
| [Fonction VAR.P](https://support.microsoft.com/office/73d1285c-108c-4843-ba5d-a51f90656f3a) | Calcule l’écart en fonction de la population entière |
| [Fonction VAR.S](https://support.microsoft.com/office/913633de-136b-449d-813e-65a00b2b990b) | Fournit une estimation de l’écart à partir d’un échantillon |
| [Fonction VARA](https://support.microsoft.com/office/3de77469-fa3a-47b4-85fd-81758a1e1d07) | Évalue la varianceen fonction d’un échantillon, y compris les nombres, le texte et les valeurs logiques |
| [Fonction VARPA](https://support.microsoft.com/office/59a62635-4e89-4fad-88ac-ce4dc0513b96) | Calcule la variance en fonction de la population entière, y compris les nombres, le texte et les valeurs logiques |
| [Fonction VDB](https://support.microsoft.com/office/dde4e207-f3fa-488d-91d2-66d55e861d73) | Renvoie l’amortissement d’un bien durant une période spécifiée ou partielle en utilisant une méthode d’amortissement dégressif |
| [Fonction VLOOKUP](https://support.microsoft.com/office/0bbc8083-26fe-4963-8ab8-93a18ad188a1) | Cherche dans la première colonne d’un tableau et se déplace horizontalement pour renvoyer la valeur d’une cellule |
| [Fonction WEEKDAY](https://support.microsoft.com/office/60e44483-2ed1-439f-8bd0-e404c190949a) | Convertit un numéro de série en jour de la semaine |
| [Fonction WEEKNUM](https://support.microsoft.com/office/e5c43a03-b4ab-426c-b411-b18c13c75340) | Convertit un numéro de série en un numéro de semaine correspondant à l’année |
| [Fonction WEIBULL.DIST](https://support.microsoft.com/office/4e783c39-9325-49be-bbc9-a83ef82b45db) | Renvoie la distribution suivant la loi de Weibull |
| [Fonction WORKDAY](https://support.microsoft.com/office/f764a5b7-05fc-4494-9486-60d494efbf33) | Renvoie le numéro de série de la date précédant ou suivant un nombre de jours ouvrés spécifié |
| [Fonction WORKDAY.INTL](https://support.microsoft.com/office/a378391c-9ba7-4678-8a39-39611a9bf81d) | Renvoie le numéro de série de la date précédant ou suivant un nombre spécifié de jours ouvrés à l’aide de paramètres indiquant le nombre de jours compris dans un week-end |
| [Fonction XIRR](https://support.microsoft.com/office/de1242ec-6477-445b-b11b-a303ad9adc9d) | Renvoie le taux de rendement interne d’une planification de flux financiers qui n’est pas nécessairement périodique |
| [Fonction XNPV](https://support.microsoft.com/office/1b42bbf6-370f-4532-a0eb-d67c16b664b7) | Renvoie la valeur actuelle nette d’une planification de flux financiers qui n’est pas nécessairement périodique |
| [Fonction XOR](https://support.microsoft.com/office/1548d4c2-5e47-4f77-9a92-0533bba14f37) | Renvoie une valeur logique exclusive OR de tous les arguments |
| [Fonction YEAR](https://support.microsoft.com/office/c64f017a-1354-490d-981f-578e8ec8d3b9) | Convertit un numéro de série en année |
| [Fonction YEARFRAC](https://support.microsoft.com/office/3844141e-c76d-4143-82b6-208454ddc6a8) | Renvoie la fraction de l’année représentant le nombre de jours entiers compris entre start_date et end_date |
| [Fonction YIELD](https://support.microsoft.com/office/f5f5ca43-c4bd-434f-8bd2-ed3c9727a4fe) | Renvoie le rendement d’un titre rapportant des intérêts périodiquement |
| [Fonction YIELDDISC](https://support.microsoft.com/office/a9dbdbae-7dae-46de-b995-615faffaaed7) | Renvoie le rendement annuel d’un titre escompté, par exemple, un bon du Trésor |
| [Fonction YIELDMAT](https://support.microsoft.com/office/ba7d1809-0d33-4bcb-96c7-6c56ec62ef6f) | Renvoie le rendement annuel d’un titre pour lequel des intérêts sont payés à l’échéance |
| [Fonction Z.TEST](https://support.microsoft.com/office/d633d5a3-2031-4614-a016-92180ad82bee) | Renvoie la valeur de probabilité unilatérale du test Z |

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Classe Functions (interface API JavaScript pour Excel)](/javascript/api/excel/excel.functions)
- [Objet Workbook Functions (interface API JavaScript pour Excel)](/javascript/api/excel/excel.workbook#excel-excel-workbook-functions-member)
