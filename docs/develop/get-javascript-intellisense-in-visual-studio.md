---
title: Obtenir JavaScript IntelliSense dans Visual Studio 2019
description: Découvrez comment utiliser JSDoc pour créer des IntelliSense pour vos variables, objets, paramètres et valeurs de retour JavaScript.
ms.date: 10/11/2019
localization_priority: Normal
ms.openlocfilehash: 6135649ce80e496d5e195b0ddb0dcb64172d41f5
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076054"
---
# <a name="get-javascript-intellisense-in-visual-studio-2019"></a><span data-ttu-id="f0396-103">Obtenir JavaScript IntelliSense dans Visual Studio 2019</span><span class="sxs-lookup"><span data-stu-id="f0396-103">Get JavaScript IntelliSense in Visual Studio 2019</span></span>

<span data-ttu-id="f0396-p101">Lorsque vous utilisez Visual Studio 2019 pour développer des compléments Office, vous pouvez utiliser JSDoc pour activer IntelliSense pour vos variables, objets, paramètres et valeurs renvoyées JavaScript. Cet article fournit une vue d’ensemble de JSDoc et explique comment vous pouvez l’utiliser pour créer IntellSense dans Visual Studio. Pour plus d’informations, voir [JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense) et [Prise en charge de JSDoc dans JavaScript](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript).</span><span class="sxs-lookup"><span data-stu-id="f0396-p101">When you use Visual Studio 2019 to develop Office Add-ins, you can use JSDoc to enable IntelliSense for your JavaScript variables, objects, parameters, and return values. This article provides an overview of JSDoc and how you can use it to create IntellSense in Visual Studio. For more details, see [JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense) and [JSDoc support in JavaScript](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript).</span></span> 

## <a name="officejs-type-definitions"></a><span data-ttu-id="f0396-107">Définitions de type Office.js</span><span class="sxs-lookup"><span data-stu-id="f0396-107">Office.js type definitions</span></span>

<span data-ttu-id="f0396-p102">Vous devez fournir les définitions des types dans Office.js pour Visual Studio. Pour ce faire, vous pouvez :</span><span class="sxs-lookup"><span data-stu-id="f0396-p102">You need to provide the definitions of the types in Office.js to Visual Studio. To do this, you can:</span></span>

- <span data-ttu-id="f0396-p103">Conserver une copie locale des fichiers Office.js dans un dossier dans votre solution nommée `\Office\1\`. Les modèles de projet Complément Office dans Visual Studio ajoutent cette copie locale lorsque vous créez un projet de complément.</span><span class="sxs-lookup"><span data-stu-id="f0396-p103">Have a local copy of the Office.js files in a folder in your solution named `\Office\1\`. The Office Add-in project templates in Visual Studio add this local copy when you create an add-in project.</span></span> 
- <span data-ttu-id="f0396-p104">Utiliser une version en ligne de Office.js en ajoutant un fichier tsconfig.json à la racine du projet d’application Web dans la solution de complément. Le fichier doit inclure le contenu suivant.</span><span class="sxs-lookup"><span data-stu-id="f0396-p104">Use an online version of Office.js by adding a tsconfig.json file to the root of the web application project in the add-in solution. The file should include the following content.</span></span>

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

## <a name="jsdoc-syntax"></a><span data-ttu-id="f0396-114">Syntaxe JSDoc</span><span class="sxs-lookup"><span data-stu-id="f0396-114">JSDoc syntax</span></span>

<span data-ttu-id="f0396-p105">La technique de base est de faire précéder la variable (ou le paramètre, etc.) d’un commentaire qui identifie son type de données. IntelliSense dans Visual Studio peut ainsi déduire ses membres. Les éléments suivants sont des exemples.</span><span class="sxs-lookup"><span data-stu-id="f0396-p105">The basic technique is to precede the variable (or parameter, and so on) with a comment that identifies its data type. This allows IntelliSense in Visual Studio to infer its members. The following are examples.</span></span>

### <a name="variable"></a><span data-ttu-id="f0396-118">Variable</span><span class="sxs-lookup"><span data-stu-id="f0396-118">Variable</span></span>

```js
/** @type {Excel.Range} */
var subsetRange;
```

![Capture d’écran montrant l IntelliSense de la variable « subsetRange ».](../images/intellisense-vs17-var.png)

### <a name="parameter"></a><span data-ttu-id="f0396-120">Paramètre</span><span class="sxs-lookup"><span data-stu-id="f0396-120">Parameter</span></span>

```js
/** @param {Word.ParagraphCollection} paragraphs */
function myFunc(paragraphs){

}
```

![Screenshot showing excerpt of IntelliSense for 'paras' parameter ('paragraphs' parameter in JavaScript example).](../images/intellisense-vs17-param.png)

### <a name="return-value"></a><span data-ttu-id="f0396-122">Valeur renvoyée</span><span class="sxs-lookup"><span data-stu-id="f0396-122">Return value</span></span>

```js
/** @returns {Word.Range} */
function myFunc() {

}
```

![Screenshot showing excerpt of IntelliSense for 'myFunc()' return value.](../images/intellisense-vs17-return.png)

### <a name="complex-types"></a><span data-ttu-id="f0396-124">Types complexes</span><span class="sxs-lookup"><span data-stu-id="f0396-124">Complex types</span></span>

```js
/** @typedef {{range: Word.Range, paragraphs: Word.ParagraphCollection}} MyType

/** @returns {MyType} */
function myFunc() {

}
```

![Capture d’IntelliSense pour la déclaration de type complexe « var myVar; » par exemple.](../images/intellisense-vs17-complex-type.png)

## <a name="see-also"></a><span data-ttu-id="f0396-126">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f0396-126">See also</span></span>

- [<span data-ttu-id="f0396-127">Développer des compléments Office avec Visual Studio</span><span class="sxs-lookup"><span data-stu-id="f0396-127">Develop Office Add-ins with Visual Studio</span></span>](develop-add-ins-visual-studio.md)
- [<span data-ttu-id="f0396-128">Déboguer des compléments Office dans Visual Studio</span><span class="sxs-lookup"><span data-stu-id="f0396-128">Debug Office Add-ins in Visual Studio</span></span>](debug-office-add-ins-in-visual-studio.md)
