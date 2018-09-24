---
ms.date: 09/20/2018
description: Découvrez les meilleures pratiques et modèles recommandés pour les fonctions personnalisées d’Excel.
title: Meilleures pratiques pour les fonctions personnalisées
ms.openlocfilehash: 1f2c0a80e62b65523fcc1673ba2ca4be444e6ce0
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/21/2018
ms.locfileid: "24068815"
---
# <a name="custom-functions-best-practices"></a><span data-ttu-id="2a190-103">Meilleures pratiques pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="2a190-103">Custom functions best practices</span></span>

<span data-ttu-id="2a190-104">Cet article décrit les meilleures pratiques pour le développement de fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="2a190-104">This article describes best practices for developing custom functions in Excel.</span></span>

## <a name="error-handling"></a><span data-ttu-id="2a190-105">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="2a190-105">Error handling</span></span>

<span data-ttu-id="2a190-106">Lorsque vous créez un complément qui définit des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution.</span><span class="sxs-lookup"><span data-stu-id="2a190-106">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="2a190-107">La gestion des erreurs pour les fonctions personnalisées est identique à la [gestion des erreurs pour l’API JavaScript d’Excel dans son ensemble](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="2a190-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="2a190-108">Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.</span><span class="sxs-lookup"><span data-stu-id="2a190-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="error-logging"></a><span data-ttu-id="2a190-109">Journalisation des erreurs</span><span class="sxs-lookup"><span data-stu-id="2a190-109">Error logging</span></span>

<span data-ttu-id="2a190-110">Vous pouvez activer la journalisation des erreurs pour votre complément de fonctions personnalisées de plusieurs façons, telles que :</span><span class="sxs-lookup"><span data-stu-id="2a190-110">You can enable error logging for your custom functions add-in in multiple ways, such as:</span></span> 

- <span data-ttu-id="2a190-111">[Utiliser la journalisation runtime](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) pour le débogage du fichier manifeste XML de votre complément.</span><span class="sxs-lookup"><span data-stu-id="2a190-111">[Use runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) to debug your add-in's XML manifest file.</span></span> 

- <span data-ttu-id="2a190-112">Utiliser des instructions `console.log` dans votre code des fonctions personnalisées pour envoyer la sortie à la console en temps réel.</span><span class="sxs-lookup"><span data-stu-id="2a190-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

> [!NOTE]
> <span data-ttu-id="2a190-113">La fonctionnalité de journalisation runtime n'est actuellement disponible que pour Office 2016 bureau.</span><span class="sxs-lookup"><span data-stu-id="2a190-113">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

## <a name="debugging"></a><span data-ttu-id="2a190-114">Débogage</span><span class="sxs-lookup"><span data-stu-id="2a190-114">Debugging</span></span>

<span data-ttu-id="2a190-115">Actuellement, la meilleure méthode pour le débogage des fonctions personnalisées Excel consiste à utiliser [Excel Online](https://www.office.com/launch/excel) et l’outil de débogage F12 natif de votre navigateur.</span><span class="sxs-lookup"><span data-stu-id="2a190-115">Currently, the best method for debugging Excel custom functions is to use [Excel Online](https://www.office.com/launch/excel) and use the F12 debugging tool native to your browser.</span></span> <span data-ttu-id="2a190-116">Des outils de débogage spécifiques pour les fonctions personnalisées pourraient être disponibles à l’avenir.</span><span class="sxs-lookup"><span data-stu-id="2a190-116">Additional debugging tools for custom functions may be available in the future.</span></span>

## <a name="mapping-names"></a><span data-ttu-id="2a190-117">Mappage de noms</span><span class="sxs-lookup"><span data-stu-id="2a190-117">Mapping names</span></span>

<span data-ttu-id="2a190-118">Par défaut, le nom d’une fonction personnalisée dans votre fichier JavaScript est déclaré généralement à l’aide de lettres toutes en majuscule et correspond exactement au nom de la fonction que l'utilisateur final voit dans Excel.</span><span class="sxs-lookup"><span data-stu-id="2a190-118">By default, the name of a custom function in your JavaScript file is typically declared using entirely uppercase letters, and corresponds exactly to the function name that end users see in Excel.</span></span> <span data-ttu-id="2a190-119">Toutefois, vous pouvez modifier ce mappage à l’aide de l'objet `CustomFunctionsMappings` pour mapper un ou plusieurs noms de fonction à partir du fichier JavaScript à des valeurs différentes que les utilisateurs finaux verront s’afficher comme noms de fonction dans Excel.</span><span class="sxs-lookup"><span data-stu-id="2a190-119">However, you can change this by using the `CustomFunctionsMappings` object to map one or more function names from the JavaScript file to different values that end users will see as function names in Excel.</span></span> <span data-ttu-id="2a190-120">Bien que vous ne soyez pas obligé d’utiliser `CustomFunctionsMapping`, il peut être utile si vous utilisez un uglifier, un webpack ou une syntaxe d'importation - qui tous ont des difficultés avec les noms de fonctions en majuscules.</span><span class="sxs-lookup"><span data-stu-id="2a190-120">Although you're not required to use `CustomFunctionsMapping`, it can be helpful if you are using an uglifier, webpack, or import syntax - all of which have difficulty with uppercase function names.</span></span>
  
<span data-ttu-id="2a190-121">L’exemple de code suivant définit une seule paire clé-valeur qui mappe le nom de la fonction JavaScript `plusFortyTwo` au nom de la fonction `ADD42` dans l’interface utilisateur d’Excel.</span><span class="sxs-lookup"><span data-stu-id="2a190-121">The following code sample defines a single key-value pair that maps the JavaScript function name `plusFortyTwo` to the `ADD42` function name in the Excel UI.</span></span> <span data-ttu-id="2a190-122">Lorsque l’utilisateur final choisit la fonction `ADD42` dans Excel, la fonction JavaScript `plusFortyTwo` s’exécute.</span><span class="sxs-lookup"><span data-stu-id="2a190-122">When the end user chooses the `ADD42` function in Excel, the `plusFortyTwo` JavaScript function will run.</span></span>

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

<span data-ttu-id="2a190-123">L’exemple de code suivant définit deux paires clé-valeur.</span><span class="sxs-lookup"><span data-stu-id="2a190-123">The following code sample defines a two key-value pairs.</span></span> <span data-ttu-id="2a190-124">La première paire mappe le nom de la fonction JavaScript `plusFifty` au nom de la fonction `ADD50` dans l’interface utilisateur d’Excel et la seconde paire mappe le nom de la fonction JavaScript `plusOneHundred` au nom de la fonction `ADD100` dans l’interface utilisateur d’Excel.</span><span class="sxs-lookup"><span data-stu-id="2a190-124">The first pair maps the JavaScript function name `plusFifty` to the `ADD50` function name in the Excel UI, and the second pair maps the JavaScript function name `plusOneHundred` to the `ADD100` function name in the Excel UI.</span></span> <span data-ttu-id="2a190-125">Lorsque l’utilisateur final choisit la fonction `ADD50` dans Excel, la fonction JavaScript `plusFifty` s’exécute.</span><span class="sxs-lookup"><span data-stu-id="2a190-125">When the end user chooses the `ADD50` function in Excel, the `plusFifty` JavaScript function will run.</span></span> <span data-ttu-id="2a190-126">Lorsque l’utilisateur final choisit la fonction `ADD100` dans Excel, la fonction JavaScript `plusOneHundred` s’exécute.</span><span class="sxs-lookup"><span data-stu-id="2a190-126">When the end user chooses the `ADD100` function in Excel, the `plusOneHundred` JavaScript function will run.</span></span>

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

 ## <a name="see-also"></a><span data-ttu-id="2a190-127">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="2a190-127">See also</span></span>

* [<span data-ttu-id="2a190-128">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="2a190-128">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="2a190-129">Métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="2a190-129">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="2a190-130">Runtime pour les fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="2a190-130">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)