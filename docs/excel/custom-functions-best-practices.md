---
ms.date: 09/20/2018
description: Découvrez les meilleures pratiques et modèles recommandés pour les fonctions personnalisées d’Excel.
title: Meilleures pratiques pour les fonctions personnalisées
ms.openlocfilehash: 3934910c397aea348c4fe2d7f95f1dc20ebeb4d3
ms.sourcegitcommit: 8ce9a8d7f41d96879c39cc5527a3007dff25bee8
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/24/2018
ms.locfileid: "24985787"
---
# <a name="custom-functions-best-practices"></a><span data-ttu-id="a557b-103">Meilleures pratiques pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a557b-103">Custom functions best practices</span></span>

<span data-ttu-id="a557b-104">Cet article décrit les meilleures pratiques pour le développement de fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="a557b-104">This article describes best practices for developing custom functions in Excel.</span></span>

## <a name="error-handling"></a><span data-ttu-id="a557b-105">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="a557b-105">Error handling</span></span>

<span data-ttu-id="a557b-106">Lorsque vous créez un complément qui définit des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution.</span><span class="sxs-lookup"><span data-stu-id="a557b-106">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="a557b-107">La gestion des erreurs pour les fonctions personnalisées est identique à la [gestion des erreurs pour l’API JavaScript d’Excel dans son ensemble](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="a557b-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="a557b-108">Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.</span><span class="sxs-lookup"><span data-stu-id="a557b-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
    let url = "https://jsonplaceholder.typicode.com/comments/" + x; 
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

## <a name="error-logging"></a><span data-ttu-id="a557b-109">Enregistrement des erreurs</span><span class="sxs-lookup"><span data-stu-id="a557b-109">Error logging</span></span>

<span data-ttu-id="a557b-110">Vous pouvez activer la journalisation des erreurs pour votre complément de fonctions personnalisées de plusieurs façons, telles que :</span><span class="sxs-lookup"><span data-stu-id="a557b-110">You can enable error logging for your custom functions add-in in multiple ways, such as:</span></span> 

- <span data-ttu-id="a557b-111">[Utilisation de la journalisation runtime](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) pour le débogage du fichier manifeste XML de votre complément.</span><span class="sxs-lookup"><span data-stu-id="a557b-111">[Use runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) to debug your add-in's XML manifest file.</span></span> 

- <span data-ttu-id="a557b-112">Utiliser des instructions `console.log` dans votre code des fonctions personnalisées pour envoyer la sortie à la console en temps réel.</span><span class="sxs-lookup"><span data-stu-id="a557b-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

> [!NOTE]
> <span data-ttu-id="a557b-113">La fonctionnalité de journalisation runtime n’est actuellement disponible que pour Office 2016 bureau.</span><span class="sxs-lookup"><span data-stu-id="a557b-113">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

## <a name="debugging"></a><span data-ttu-id="a557b-114">Débogage</span><span class="sxs-lookup"><span data-stu-id="a557b-114">Debugging</span></span>

<span data-ttu-id="a557b-115">Actuellement, la meilleure méthode pour le débogage des fonctions personnalisées Excel consiste à premier [sideload](../testing/sideload-office-add-ins-for-testing.md) votre complément dans Excel Online.</span><span class="sxs-lookup"><span data-stu-id="a557b-115">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within Excel Online.</span></span> <span data-ttu-id="a557b-116">Ensuite, vous pouvez déboguer vos fonctions personnalisées à l’aide de l' [outil de débogage F12 natif de votre navigateur](../testing/debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="a557b-116">Then you can debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md).</span></span>

<span data-ttu-id="a557b-117">Si votre complément ne parvient pas à s’enregistrer, [vérifiez que les certificats SSL sont correctement configurés](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour le serveur web qui héberge votre application de complément.</span><span class="sxs-lookup"><span data-stu-id="a557b-117">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

## <a name="mapping-names"></a><span data-ttu-id="a557b-118">Mappage de noms</span><span class="sxs-lookup"><span data-stu-id="a557b-118">Mapping names</span></span>

<span data-ttu-id="a557b-119">Par défaut, le nom d’une fonction personnalisée dans votre fichier JavaScript est déclaré généralement à l’aide de lettres toutes en majuscule et correspond exactement au nom de la fonction que l'utilisateur final voit dans Excel.</span><span class="sxs-lookup"><span data-stu-id="a557b-119">By default, the name of a custom function in your JavaScript file is typically declared using entirely uppercase letters, and corresponds exactly to the function name that end users see in Excel.</span></span> <span data-ttu-id="a557b-120">Toutefois, vous pouvez modifier ce mappage à l’aide de l'objet `CustomFunctionsMappings` pour mapper un ou plusieurs noms de fonction à partir du fichier JavaScript à des valeurs différentes que les utilisateurs finaux verront s’afficher comme noms de fonction dans Excel.</span><span class="sxs-lookup"><span data-stu-id="a557b-120">However, you can change this by using the `CustomFunctionsMappings` object to map one or more function names from the JavaScript file to different values that end users will see as function names in Excel.</span></span> <span data-ttu-id="a557b-121">Cela peut être utile si vous utilisez un uglifier, un webpack ou une syntaxe d’importation - qui tous ont des difficultés avec les noms de fonctions en majuscules.</span><span class="sxs-lookup"><span data-stu-id="a557b-121">Although you're not required to use , it can be helpful if you are using an uglifier, webpack, or import syntax - all of which have difficulty with uppercase function names.</span></span> <span data-ttu-id="a557b-122">`CustomFunctionsMappings` Il est éventuellement facultatif pour les projets utilisant JavaScript, mais vous devez vous en servir si votre projet utilise des caractères dactylographiés.</span><span class="sxs-lookup"><span data-stu-id="a557b-122">`CustomFunctionsMappings` is possibly optional for projects using JavaScript but must be used if your project uses TypeScript.</span></span>  
  
<span data-ttu-id="a557b-123">L’exemple de code suivant définit une seule paire clé-valeur qui mappe le nom de la fonction JavaScript `plusFortyTwo` au nom de la fonction `ADD42` dans l’interface utilisateur d’Excel.</span><span class="sxs-lookup"><span data-stu-id="a557b-123">The following code sample defines a single key-value pair that maps the JavaScript function name `plusFortyTwo` to the `ADD42` function name in the Excel UI.</span></span> <span data-ttu-id="a557b-124">Lorsque l’utilisateur final choisit la fonction `ADD42` dans Excel, la fonction JavaScript `plusFortyTwo` s’exécute.</span><span class="sxs-lookup"><span data-stu-id="a557b-124">When the end user chooses the `ADD42` function in Excel, the `plusFortyTwo` JavaScript function will run.</span></span>

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

<span data-ttu-id="a557b-125">L’exemple de code suivant définit deux paires clé-valeur.</span><span class="sxs-lookup"><span data-stu-id="a557b-125">The following code sample defines a two key-value pairs.</span></span> <span data-ttu-id="a557b-126">La première paire mappe le nom de la fonction JavaScript `plusFifty` au nom de la fonction `ADD50` dans l’interface utilisateur d’Excel et la seconde paire mappe le nom de la fonction JavaScript `plusOneHundred` au nom de la fonction `ADD100` dans l’interface utilisateur d’Excel.</span><span class="sxs-lookup"><span data-stu-id="a557b-126">The first pair maps the JavaScript function name `plusFifty` to the `ADD50` function name in the Excel UI, and the second pair maps the JavaScript function name `plusOneHundred` to the `ADD100` function name in the Excel UI.</span></span> <span data-ttu-id="a557b-127">Lorsque l’utilisateur final choisit la fonction `ADD50` dans Excel, la fonction JavaScript `plusFifty` s’exécute.</span><span class="sxs-lookup"><span data-stu-id="a557b-127">When the end user chooses the `ADD50` function in Excel, the `plusFifty` JavaScript function will run.</span></span> <span data-ttu-id="a557b-128">Lorsque l’utilisateur final choisit la fonction `ADD100` dans Excel, la fonction JavaScript `plusOneHundred` s’exécute.</span><span class="sxs-lookup"><span data-stu-id="a557b-128">When the end user chooses the `ADD100` function in Excel, the `plusOneHundred` JavaScript function will run.</span></span>

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

 ## <a name="see-also"></a><span data-ttu-id="a557b-129">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a557b-129">See also</span></span>

* [<span data-ttu-id="a557b-130">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="a557b-130">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="a557b-131">Métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a557b-131">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="a557b-132">Runtime pour les fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="a557b-132">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
