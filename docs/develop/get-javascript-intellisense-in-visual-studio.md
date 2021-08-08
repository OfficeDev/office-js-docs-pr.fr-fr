---
title: Obtenir JavaScript IntelliSense dans Visual Studio 2019
description: Découvrez comment utiliser JSDoc pour créer des IntelliSense pour vos variables, objets, paramètres et valeurs de retour JavaScript.
ms.date: 10/11/2019
localization_priority: Normal
ms.openlocfilehash: 35c3563573cb9edfbe975504b4b04330bbd1a0422c5141ee9f4141db7a478807
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57080657"
---
# <a name="get-javascript-intellisense-in-visual-studio-2019"></a>Obtenir JavaScript IntelliSense dans Visual Studio 2019

Lorsque vous utilisez Visual Studio 2019 pour développer des compléments Office, vous pouvez utiliser JSDoc pour activer IntelliSense pour vos variables, objets, paramètres et valeurs renvoyées JavaScript. Cet article fournit une vue d’ensemble de JSDoc et explique comment vous pouvez l’utiliser pour créer IntellSense dans Visual Studio. Pour plus d’informations, voir [JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense) et [Prise en charge de JSDoc dans JavaScript](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript). 

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

![Screenshot showing excerpt of IntelliSense for 'subsetRange' variable.](../images/intellisense-vs17-var.png)

### <a name="parameter"></a>Paramètre

```js
/** @param {Word.ParagraphCollection} paragraphs */
function myFunc(paragraphs){

}
```

![Screenshot showing excerpt of IntelliSense for 'paras' parameter ('paragraphs' parameter in JavaScript example).](../images/intellisense-vs17-param.png)

### <a name="return-value"></a>Valeur renvoyée

```js
/** @returns {Word.Range} */
function myFunc() {

}
```

![Screenshot showing excerpt of IntelliSense for 'myFunc()' return value.](../images/intellisense-vs17-return.png)

### <a name="complex-types"></a>Types complexes

```js
/** @typedef {{range: Word.Range, paragraphs: Word.ParagraphCollection}} MyType

/** @returns {MyType} */
function myFunc() {

}
```

![Capture d’IntelliSense pour la déclaration de type complexe « var myVar; » par exemple.](../images/intellisense-vs17-complex-type.png)

## <a name="see-also"></a>Voir aussi

- [Développer des compléments Office avec Visual Studio](develop-add-ins-visual-studio.md)
- [Déboguer des compléments Office dans Visual Studio](debug-office-add-ins-in-visual-studio.md)
