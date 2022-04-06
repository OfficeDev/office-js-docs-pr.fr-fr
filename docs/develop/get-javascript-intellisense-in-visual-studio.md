---
title: Obtenir JavaScript IntelliSense dans Visual Studio
description: Découvrez comment utiliser JSDoc pour créer IntelliSense pour vos variables, objets, paramètres et valeurs de retour JavaScript.
ms.date: 04/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4fb18199c53b7b836079754b5fb4561b76a968b1
ms.sourcegitcommit: 3c5ede9c4f9782947cea07646764f76156504ff9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/06/2022
ms.locfileid: "64682258"
---
# <a name="get-javascript-intellisense-in-visual-studio"></a>Obtenir JavaScript IntelliSense dans Visual Studio

Lorsque vous utilisez Visual Studio 2019 et versions ultérieures pour développer des compléments Office, vous pouvez utiliser JSDoc pour activer IntelliSense pour vos variables, objets, paramètres et valeurs de retour JavaScript. Cet article fournit une vue d’ensemble de JSDoc et explique comment vous pouvez l’utiliser pour créer IntellSense dans Visual Studio. Pour plus d’informations, voir [JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense) et [Prise en charge de JSDoc dans JavaScript](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript). 

## <a name="officejs-type-definitions"></a>Définitions de type Office.js

Vous devez fournir les définitions des types dans Office.js pour Visual Studio. Pour ce faire, vous pouvez :

- Conserver une copie locale des fichiers Office.js dans un dossier dans votre solution nommée `\Office\1\`. Les modèles de projet Complément Office dans Visual Studio ajoutent cette copie locale lorsque vous créez un projet de complément. 
- Utiliser une version en ligne de Office.js en ajoutant un fichier tsconfig.json à la racine du projet d’application Web dans la solution de complément. Le fichier doit inclure le contenu suivant.

    ```json
        {
            "compilerOptions": {
                "allowJs": true,            // These settings apply to JavaScript files also.
                "noEmit":  true             // Do not compile the JS (or TS) files in this project.
            },
            "exclude": [
                "node_modules",             // Don't include any JavaScript found under "node_modules".
                "Scripts/Office/1"          // Suppress loading all the JavaScript files from the Office NuGet package.
            ],
            "typeAcquisition": {
                "enable": true,             // Enable automatic fetching of type definitions for detected JavaScript libraries.
                "include": [ "office-js" ]  // Ensure that the "Office-js" type definition is fetched.
            }
        }
    ```

## <a name="jsdoc-syntax"></a>Syntaxe JSDoc

La technique de base est de faire précéder la variable (ou le paramètre, etc.) d’un commentaire qui identifie son type de données. IntelliSense dans Visual Studio peut ainsi déduire ses membres. Les éléments suivants sont des exemples.

### <a name="variable"></a>Variable

```js
/** @type {Excel.Range} */
var subsetRange;
```

![Capture d’écran montrant un extrait d’IntelliSense pour la variable « subsetRange ».](../images/intellisense-vs17-var.png)

### <a name="parameter"></a>Parameter

```js
/** @param {Word.ParagraphCollection} paragraphs */
function myFunc(paragraphs){

}
```

![Capture d’écran montrant un extrait d’IntelliSense pour le paramètre « paras » (paramètre « paragraphs » dans l’exemple JavaScript).](../images/intellisense-vs17-param.png)

### <a name="return-value"></a>Valeur renvoyée

```js
/** @returns {Word.Range} */
function myFunc() {

}
```

![Capture d’écran montrant un extrait d’IntelliSense pour la valeur de retour « myFunc() ».](../images/intellisense-vs17-return.png)

### <a name="complex-types"></a>Types complexes

```js
/** @typedef {{range: Word.Range, paragraphs: Word.ParagraphCollection}} MyType

/** @returns {MyType} */
function myFunc() {

}
```

![Capture d’écran montrant IntelliSense pour la déclaration de type complexe « var myVar; », par exemple.](../images/intellisense-vs17-complex-type.png)

## <a name="see-also"></a>Voir aussi

- [Développer des compléments Office avec Visual Studio](develop-add-ins-visual-studio.md)
- [Déboguer des compléments Office dans Visual Studio](debug-office-add-ins-in-visual-studio.md)
