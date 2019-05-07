---
ms.date: 05/03/2019
description: Découvrez les conditions requises pour les noms des fonctions personnalisées Excel et éviter les pièges de dénomination courants.
title: Instructions d’affectation de noms pour les fonctions personnalisées dans Excel
localization_priority: Normal
ms.openlocfilehash: 3abe04eebfa703666b70ecbde1c68ab0c942003c
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628045"
---
# <a name="naming-guidelines"></a><span data-ttu-id="91f83-103">Instructions d’affectation de noms</span><span class="sxs-lookup"><span data-stu-id="91f83-103">Naming guidelines</span></span>

<span data-ttu-id="91f83-104">Une fonction personnalisée est identifiée par un **ID** et une propriété de **nom** dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="91f83-104">A custom function is identified by an **id** and **name** property in the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

- <span data-ttu-id="91f83-105">La fonction `id` est utilisée pour identifier des fonctions personnalisées de manière unique dans votre code JavaScript.</span><span class="sxs-lookup"><span data-stu-id="91f83-105">The function `id` is used to uniquely identify custom functions in your JavaScript code.</span></span> 
- <span data-ttu-id="91f83-106">La fonction `name` est utilisée comme nom complet qui s’affiche pour un utilisateur dans Excel.</span><span class="sxs-lookup"><span data-stu-id="91f83-106">The function `name` is used as the display name that appears to a user in Excel.</span></span> 

<span data-ttu-id="91f83-107">Une fonction `name` peut différer de la `id`fonction, par exemple à des fins de localisation.</span><span class="sxs-lookup"><span data-stu-id="91f83-107">A function `name` can differ from the function `id`, such as for localization purposes.</span></span> <span data-ttu-id="91f83-108">En règle générale, les fonctions `name` d’une fonction doivent rester les `id` mêmes que s’il n’y a aucune raison impérieuse de les différencier.</span><span class="sxs-lookup"><span data-stu-id="91f83-108">In general, a function's `name` should stay the same as the `id` if there is no compelling reason for them to differ.</span></span>

<span data-ttu-id="91f83-109">Une fonction `name` et `id` partagent des exigences communes:</span><span class="sxs-lookup"><span data-stu-id="91f83-109">A function's `name` and `id` share some common requirements:</span></span>

- <span data-ttu-id="91f83-110">Une fonction `id` ne peut utiliser que les caractères A à Z, les chiffres 0 à 9, les traits de soulignement et les points.</span><span class="sxs-lookup"><span data-stu-id="91f83-110">A function's `id` may only use characters A through Z, numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="91f83-111">Une fonction `name` peut utiliser n’importe quel caractère alphabétique Unicode, des traits de soulignement et des points.</span><span class="sxs-lookup"><span data-stu-id="91f83-111">A function's `name` may use any Unicode alphabetic characters, underscores, and periods.</span></span>

- <span data-ttu-id="91f83-112">Les deux `name` fonctions `id` et doivent commencer par une lettre et comporter une limite minimale de trois caractères.</span><span class="sxs-lookup"><span data-stu-id="91f83-112">Both function `name` and `id` must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="91f83-113">Excel utilise des lettres majuscules pour les noms de fonctions intégrées ( `SUM`par exemple,).</span><span class="sxs-lookup"><span data-stu-id="91f83-113">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="91f83-114">Par conséquent, envisagez d’utiliser des lettres majuscules `id` pour votre fonction personnalisée et constitue `name` une meilleure pratique.</span><span class="sxs-lookup"><span data-stu-id="91f83-114">Therefore, consider using uppercase letters for your custom function's `name` and `id` as a best practice.</span></span>

<span data-ttu-id="91f83-115">Une fonction `name` ne doit pas être nommée de la manière suivante:</span><span class="sxs-lookup"><span data-stu-id="91f83-115">A function's `name` shouldn't be named the same as:</span></span>

- <span data-ttu-id="91f83-116">Toutes les cellules comprises entre a1 et XFD1048576 ou toutes les cellules comprises entre R1C1 et R1048576C16384.</span><span class="sxs-lookup"><span data-stu-id="91f83-116">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="91f83-117">N’importe quelle fonction macro Excel 4,0 ( `RUN`telle `ECHO`que,).</span><span class="sxs-lookup"><span data-stu-id="91f83-117">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="91f83-118">Pour obtenir une liste complète de ces fonctions, consultez [cet article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).</span><span class="sxs-lookup"><span data-stu-id="91f83-118">For a full list of these functions, see [this article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="91f83-119">Conflits de noms</span><span class="sxs-lookup"><span data-stu-id="91f83-119">Naming conflicts</span></span>

<span data-ttu-id="91f83-120">Si votre fonction `name` est identique à une fonction `name` dans un complément qui existe déjà, le **#REF!**</span><span class="sxs-lookup"><span data-stu-id="91f83-120">If your function `name` is the same as a function `name` in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="91f83-121">une erreur apparaît dans votre classeur.</span><span class="sxs-lookup"><span data-stu-id="91f83-121">error will appear in your workbook.</span></span>

<span data-ttu-id="91f83-122">Pour résoudre un conflit d’affectation de noms `name` , modifiez le dans votre complément et renouvelez la fonction.</span><span class="sxs-lookup"><span data-stu-id="91f83-122">To fix a naming conflict, change the `name` in your add-in and try the function again.</span></span> <span data-ttu-id="91f83-123">Vous pouvez également désinstaller le complément avec le nom conflictuel.</span><span class="sxs-lookup"><span data-stu-id="91f83-123">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="91f83-124">Ou, si vous testez votre complément dans différents environnements, essayez d’utiliser un espace de noms différent pour différencier votre fonction ( `NAMESPACE_NAMEOFFUNCTION`telle que).</span><span class="sxs-lookup"><span data-stu-id="91f83-124">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as `NAMESPACE_NAMEOFFUNCTION`).</span></span>

## <a name="best-practices"></a><span data-ttu-id="91f83-125">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="91f83-125">Best practices</span></span>

- <span data-ttu-id="91f83-126">Envisagez d’ajouter plusieurs arguments à une fonction plutôt que de créer plusieurs fonctions avec des noms identiques ou similaires.</span><span class="sxs-lookup"><span data-stu-id="91f83-126">Consider adding multiple arguments to a function rather than creating multiple functions with the same or similar names.</span></span>
- <span data-ttu-id="91f83-127">Les noms de fonction doivent indiquer l’action de la fonction, `=GETZIPCODE` par exemple `ZIPCODE`au lieu de.</span><span class="sxs-lookup"><span data-stu-id="91f83-127">Function names should indicate the action of the function, such as `=GETZIPCODE` instead of `ZIPCODE`.</span></span>
- <span data-ttu-id="91f83-128">Évitez les abréviations ambiguës dans les noms de fonction.</span><span class="sxs-lookup"><span data-stu-id="91f83-128">Avoid ambiguous abbreviations in function names.</span></span> <span data-ttu-id="91f83-129">La clarté est plus importante que la concision.</span><span class="sxs-lookup"><span data-stu-id="91f83-129">Clarity is more important than brevity.</span></span> <span data-ttu-id="91f83-130">Choisissez un nom tel `=INCREASETIME` que plutôt `=INC`que.</span><span class="sxs-lookup"><span data-stu-id="91f83-130">Choose a name like `=INCREASETIME` rather than `=INC`.</span></span>
- <span data-ttu-id="91f83-131">Utilisez régulièrement les mêmes verbes pour les fonctions qui effectuent des actions similaires.</span><span class="sxs-lookup"><span data-stu-id="91f83-131">Consistently use the same verbs for functions which perform similar actions.</span></span> <span data-ttu-id="91f83-132">Par exemple, utilisez `=DELETEZIPCODE` and `=DELETEADDRESS`, et non `=DELETEZIPCODE` et `=REMOVEADDRESS`.</span><span class="sxs-lookup"><span data-stu-id="91f83-132">For example, use `=DELETEZIPCODE` and `=DELETEADDRESS`, rather than `=DELETEZIPCODE` and `=REMOVEADDRESS`.</span></span>

## <a name="localizing-function-names"></a><span data-ttu-id="91f83-133">Localisation des noms de fonction</span><span class="sxs-lookup"><span data-stu-id="91f83-133">Localizing function names</span></span>

<span data-ttu-id="91f83-134">Vous pouvez localiser vos noms de fonction pour différentes langues à l’aide de fichiers JSON distincts et remplacer les valeurs dans le fichier manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="91f83-134">You can localize your function names for different languages using separate JSON files and override values in your add-in's manifest file.</span></span> <span data-ttu-id="91f83-135">Nous vous recommandons de ne pas donner à vos fonctions `id` une `name` ou une fonction Excel intégrée dans un autre langage, car cela peut entraîner des conflits avec des fonctions localisées.</span><span class="sxs-lookup"><span data-stu-id="91f83-135">As a best practice, avoid giving your functions an `id` or `name` that is a built-in Excel function in another language as this could conflict with localized functions.</span></span>

<span data-ttu-id="91f83-136">Pour obtenir des informations complètes sur la localisation, voir [Localize Custom Functions](custom-functions-localize.md)</span><span class="sxs-lookup"><span data-stu-id="91f83-136">For full information on localizing, see [Localize custom functions](custom-functions-localize.md)</span></span>

## <a name="next-steps"></a><span data-ttu-id="91f83-137">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="91f83-137">Next steps</span></span>
<span data-ttu-id="91f83-138">Découvrez les [meilleures pratiques en matière de gestion des erreurs](custom-functions-errors.md).</span><span class="sxs-lookup"><span data-stu-id="91f83-138">Learn about [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="91f83-139">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="91f83-139">See also</span></span>

* [<span data-ttu-id="91f83-140">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="91f83-140">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="91f83-141">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="91f83-141">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="91f83-142">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="91f83-142">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="91f83-143">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="91f83-143">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
