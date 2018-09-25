---
ms.date: 09/20/2018
description: Découvrez les meilleures pratiques et modèles recommandés pour les fonctions personnalisées d’Excel.
title: Meilleures pratiques pour les fonctions personnalisées
ms.openlocfilehash: 4fe0ddc36ce1b08ea360bb556121e76cd57c3823
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004909"
---
# <a name="custom-functions-best-practices"></a><span data-ttu-id="c32b7-103">Meilleures pratiques pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="c32b7-103">Custom functions best practices</span></span>

<span data-ttu-id="c32b7-104">Cet article décrit les meilleures pratiques pour le développement de fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="c32b7-104">This article describes best practices for developing custom functions in Excel.</span></span>

## <a name="error-handling"></a><span data-ttu-id="c32b7-105">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="c32b7-105">Error handling</span></span>

<span data-ttu-id="c32b7-106">Lorsque vous créez un complément qui définit des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution.</span><span class="sxs-lookup"><span data-stu-id="c32b7-106">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="c32b7-107">La gestion des erreurs pour les fonctions personnalisées est identique à la [gestion des erreurs pour l’API JavaScript d’Excel dans son ensemble](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="c32b7-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="c32b7-108">Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.</span><span class="sxs-lookup"><span data-stu-id="c32b7-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
    let url = "https://yourhypotheticalapi.com/comments/" + x; 
    return fetch(url)
        .then(function (data) {
            return data.json();
        })
        .then((json) => {
            return json.body;
        })
        .catch(function (error) {
            throw error;
        })
}
```

## <a name="debugging"></a><span data-ttu-id="c32b7-109">Débogage</span><span class="sxs-lookup"><span data-stu-id="c32b7-109">Debugging</span></span>
<span data-ttu-id="c32b7-110">Actuellement, la meilleure méthode pour le débogage des fonctions personnalisées Excel consiste à premier [sideload](../testing/sideload-office-add-ins-for-testing.md) votre complément dans **Excel Online**.</span><span class="sxs-lookup"><span data-stu-id="c32b7-110">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**.</span></span> <span data-ttu-id="c32b7-111">Ensuite, vous pouvez déboguer vos fonctions personnalisées à l’aide de l’[outil de débogage F12 natif de votre navigateur](../testing/debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="c32b7-111">You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md).</span></span> <span data-ttu-id="c32b7-112">Utiliser des instructions `console.log` dans votre code des fonctions personnalisées pour envoyer la sortie à la console en temps réel.</span><span class="sxs-lookup"><span data-stu-id="c32b7-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

<span data-ttu-id="c32b7-113">Si votre complément ne parvient pas à s’enregistrer, [vérifiez que les certificats SSL sont correctement configurés](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour le serveur web qui héberge votre application de complément.</span><span class="sxs-lookup"><span data-stu-id="c32b7-113">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

<span data-ttu-id="c32b7-114">Si vous testez votre complément dans Office 2016 bureau, vous pouvez activer la [journalisation runtime](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) pour résoudre les problèmes du fichier manifeste XML de votre complément, ainsi que plusieurs conditions d’installation et d’exécution.</span><span class="sxs-lookup"><span data-stu-id="c32b7-114">If you are testing your add-in in Office 2016 desktop, you can enable [runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) to debug issues with your add-in's XML manifest file as well as several installation and runtime conditions.</span></span> 


## <a name="mapping-names"></a><span data-ttu-id="c32b7-115">Mappage de noms</span><span class="sxs-lookup"><span data-stu-id="c32b7-115">Mapping names</span></span>

<span data-ttu-id="c32b7-116">Par défaut, le nom d’une fonction personnalisée dans votre fichier JavaScript est déclaré généralement à l’aide de lettres toutes en majuscule et correspond exactement au nom de la fonction que l'utilisateur final voit dans Excel.</span><span class="sxs-lookup"><span data-stu-id="c32b7-116">By default, the name of a custom function in your JavaScript file is typically declared using entirely uppercase letters, and corresponds exactly to the function name that end users see in Excel.</span></span> <span data-ttu-id="c32b7-117">Toutefois, vous pouvez modifier ce mappage à l’aide de l'objet `CustomFunctionsMappings` pour mapper un ou plusieurs noms de fonction à partir du fichier JavaScript à des valeurs différentes que les utilisateurs finaux verront s’afficher comme noms de fonction dans Excel.</span><span class="sxs-lookup"><span data-stu-id="c32b7-117">However, you can change this by using the `CustomFunctionsMappings` object to map one or more function names from the JavaScript file to different values that end users will see as function names in Excel.</span></span> <span data-ttu-id="c32b7-118">Cela peut être utile si vous utilisez un uglifier, un webpack ou une syntaxe d’importation - qui ont tous des difficultés avec les noms de fonctions en majuscules.</span><span class="sxs-lookup"><span data-stu-id="c32b7-118">Although you're not required to use , it can be helpful if you are using an uglifier, webpack, or import syntax - all of which have difficulty with uppercase function names.</span></span> <span data-ttu-id="c32b7-119">`CustomFunctionsMappings` Il est éventuellement facultatif pour les projets utilisant JavaScript, mais vous devez vous en servir si votre projet utilise des caractères dactylographiés.</span><span class="sxs-lookup"><span data-stu-id="c32b7-119">`CustomFunctionsMappings` is possibly optional for projects using JavaScript but must be used if your project uses TypeScript.</span></span>  
  
<span data-ttu-id="c32b7-120">L’exemple de code suivant définit une seule paire clé-valeur qui mappe le nom de la fonction JavaScript `plusFortyTwo` au nom de la fonction `ADD42` dans l’interface utilisateur d’Excel.</span><span class="sxs-lookup"><span data-stu-id="c32b7-120">The following code sample defines a single key-value pair that maps the JavaScript function name `plusFortyTwo` to the `ADD42` function name in the Excel UI.</span></span> <span data-ttu-id="c32b7-121">Lorsque l’utilisateur final choisit la fonction `ADD42` dans Excel, la fonction JavaScript `plusFortyTwo` s’exécute.</span><span class="sxs-lookup"><span data-stu-id="c32b7-121">When the end user chooses the `ADD42` function in Excel, the `plusFortyTwo` JavaScript function will run.</span></span>

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

<span data-ttu-id="c32b7-122">L’exemple de code suivant définit deux paires clé-valeur.</span><span class="sxs-lookup"><span data-stu-id="c32b7-122">The following code sample defines a two key-value pairs.</span></span> <span data-ttu-id="c32b7-123">La première paire mappe le nom de la fonction JavaScript `plusFifty` au nom de la fonction `ADD50` dans l’interface utilisateur d’Excel et la seconde paire mappe le nom de la fonction JavaScript `plusOneHundred` au nom de la fonction `ADD100` dans l’interface utilisateur d’Excel.</span><span class="sxs-lookup"><span data-stu-id="c32b7-123">The first pair maps the JavaScript function name `plusFifty` to the `ADD50` function name in the Excel UI, and the second pair maps the JavaScript function name `plusOneHundred` to the `ADD100` function name in the Excel UI.</span></span> <span data-ttu-id="c32b7-124">Lorsque l’utilisateur final choisit la fonction `ADD50` dans Excel, la fonction JavaScript `plusFifty` s’exécute.</span><span class="sxs-lookup"><span data-stu-id="c32b7-124">When the end user chooses the `ADD50` function in Excel, the `plusFifty` JavaScript function will run.</span></span> <span data-ttu-id="c32b7-125">Lorsque l’utilisateur final choisit la fonction `ADD100` dans Excel, la fonction JavaScript `plusOneHundred` s’exécute.</span><span class="sxs-lookup"><span data-stu-id="c32b7-125">When the end user chooses the `ADD100` function in Excel, the `plusOneHundred` JavaScript function will run.</span></span>

```js
function plusFifty(num) {
    return num + 50;  
} 

function plusOneHundred(num) {
    return num + 100;  
}  
  
CustomFunctionsMappings = {
    "plusFifty" : ADD50,  
    "plusOneHundred" : ADD100
}
 ```

 ## <a name="see-also"></a><span data-ttu-id="c32b7-126">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c32b7-126">See also</span></span>

- [<span data-ttu-id="c32b7-127">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="c32b7-127">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
- [<span data-ttu-id="c32b7-128">Métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="c32b7-128">Custom functions metadata</span></span>](custom-functions-json.md)
- [<span data-ttu-id="c32b7-129">Runtime pour les fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="c32b7-129">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
