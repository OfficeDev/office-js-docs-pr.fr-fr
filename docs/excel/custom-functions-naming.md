---
ms.date: 12/28/2019
description: Découvrez les conditions requises pour les noms des fonctions personnalisées Excel et éviter les pièges de dénomination courants.
title: Instructions d’affectation de noms pour les fonctions personnalisées dans Excel
localization_priority: Normal
ms.openlocfilehash: 547d1edd1a8103afce8b233f77085296cd85642b
ms.sourcegitcommit: abe8188684b55710261c69e206de83d3a6bd2ed3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/08/2020
ms.locfileid: "40969513"
---
# <a name="naming-guidelines"></a><span data-ttu-id="b7180-103">Instructions d’attribution de noms</span><span class="sxs-lookup"><span data-stu-id="b7180-103">Naming guidelines</span></span>

<span data-ttu-id="b7180-104">Une fonction personnalisée est identifiée par un **ID** et une propriété de **nom** dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="b7180-104">A custom function is identified by an **id** and **name** property in the JSON metadata file.</span></span>

- <span data-ttu-id="b7180-105">La fonction `id` est utilisée pour identifier des fonctions personnalisées de manière unique dans votre code JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b7180-105">The function `id` is used to uniquely identify custom functions in your JavaScript code.</span></span>
- <span data-ttu-id="b7180-106">La fonction `name` est utilisée comme nom complet qui s’affiche pour un utilisateur dans Excel.</span><span class="sxs-lookup"><span data-stu-id="b7180-106">The function `name` is used as the display name that appears to a user in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="b7180-107">Une fonction `name` peut différer de la `id`fonction, par exemple à des fins de localisation.</span><span class="sxs-lookup"><span data-stu-id="b7180-107">A function `name` can differ from the function `id`, such as for localization purposes.</span></span> <span data-ttu-id="b7180-108">En règle générale, les fonctions `name` d’une fonction doivent rester les `id` mêmes que s’il n’y a aucune raison impérieuse de les différencier.</span><span class="sxs-lookup"><span data-stu-id="b7180-108">In general, a function's `name` should stay the same as the `id` if there is no compelling reason for them to differ.</span></span>

<span data-ttu-id="b7180-109">Une fonction `name` et `id` partagent des exigences communes :</span><span class="sxs-lookup"><span data-stu-id="b7180-109">A function's `name` and `id` share some common requirements:</span></span>

- <span data-ttu-id="b7180-110">Une fonction `id` ne peut utiliser que les caractères A à Z, les chiffres 0 à 9, les traits de soulignement et les points.</span><span class="sxs-lookup"><span data-stu-id="b7180-110">A function's `id` may only use characters A through Z, numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="b7180-111">Une fonction `name` peut utiliser n’importe quel caractère alphabétique Unicode, des traits de soulignement et des points.</span><span class="sxs-lookup"><span data-stu-id="b7180-111">A function's `name` may use any Unicode alphabetic characters, underscores, and periods.</span></span>

- <span data-ttu-id="b7180-112">Les deux `name` fonctions `id` et doivent commencer par une lettre et comporter une limite minimale de trois caractères.</span><span class="sxs-lookup"><span data-stu-id="b7180-112">Both function `name` and `id` must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="b7180-113">Excel utilise des lettres majuscules pour les noms de fonctions intégrées ( `SUM`par exemple,).</span><span class="sxs-lookup"><span data-stu-id="b7180-113">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="b7180-114">Par conséquent, envisagez d’utiliser des lettres majuscules `id` pour votre fonction personnalisée et constitue `name` une meilleure pratique.</span><span class="sxs-lookup"><span data-stu-id="b7180-114">Therefore, consider using uppercase letters for your custom function's `name` and `id` as a best practice.</span></span>

<span data-ttu-id="b7180-115">Une fonction `name` ne doit pas être nommée de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="b7180-115">A function's `name` shouldn't be named the same as:</span></span>

- <span data-ttu-id="b7180-116">Toutes les cellules comprises entre a1 et XFD1048576 ou toutes les cellules comprises entre R1C1 et R1048576C16384.</span><span class="sxs-lookup"><span data-stu-id="b7180-116">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="b7180-117">N’importe quelle fonction macro Excel 4,0 ( `RUN`telle `ECHO`que,).</span><span class="sxs-lookup"><span data-stu-id="b7180-117">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="b7180-118">Pour obtenir la liste complète de ces fonctions, consultez [le document de référence des fonctions de macro Excel](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf).</span><span class="sxs-lookup"><span data-stu-id="b7180-118">For a full list of these functions, see [this Excel Macro Functions Reference document](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="b7180-119">Conflits de noms</span><span class="sxs-lookup"><span data-stu-id="b7180-119">Naming conflicts</span></span>

<span data-ttu-id="b7180-120">Si votre fonction `name` est identique à une fonction `name` dans un complément qui existe déjà, le **#REF !**</span><span class="sxs-lookup"><span data-stu-id="b7180-120">If your function `name` is the same as a function `name` in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="b7180-121">une erreur apparaît dans votre classeur.</span><span class="sxs-lookup"><span data-stu-id="b7180-121">error will appear in your workbook.</span></span>

<span data-ttu-id="b7180-122">Pour résoudre un conflit d’affectation de noms `name` , modifiez le dans votre complément et renouvelez la fonction.</span><span class="sxs-lookup"><span data-stu-id="b7180-122">To fix a naming conflict, change the `name` in your add-in and try the function again.</span></span> <span data-ttu-id="b7180-123">Vous pouvez également désinstaller le complément avec le nom conflictuel.</span><span class="sxs-lookup"><span data-stu-id="b7180-123">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="b7180-124">Ou, si vous testez votre complément dans différents environnements, essayez d’utiliser un espace de noms différent pour différencier votre fonction ( `NAMESPACE_NAMEOFFUNCTION`telle que).</span><span class="sxs-lookup"><span data-stu-id="b7180-124">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as `NAMESPACE_NAMEOFFUNCTION`).</span></span>

## <a name="best-practices"></a><span data-ttu-id="b7180-125">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="b7180-125">Best practices</span></span>

- <span data-ttu-id="b7180-126">Envisagez d’ajouter plusieurs arguments à une fonction plutôt que de créer plusieurs fonctions avec des noms identiques ou similaires.</span><span class="sxs-lookup"><span data-stu-id="b7180-126">Consider adding multiple arguments to a function rather than creating multiple functions with the same or similar names.</span></span>
- <span data-ttu-id="b7180-127">Les noms de fonction doivent indiquer l’action de la fonction, `=GETZIPCODE` par exemple `ZIPCODE`au lieu de.</span><span class="sxs-lookup"><span data-stu-id="b7180-127">Function names should indicate the action of the function, such as `=GETZIPCODE` instead of `ZIPCODE`.</span></span>
- <span data-ttu-id="b7180-128">Évitez les abréviations ambiguës dans les noms de fonction.</span><span class="sxs-lookup"><span data-stu-id="b7180-128">Avoid ambiguous abbreviations in function names.</span></span> <span data-ttu-id="b7180-129">La clarté est plus importante que la concision.</span><span class="sxs-lookup"><span data-stu-id="b7180-129">Clarity is more important than brevity.</span></span> <span data-ttu-id="b7180-130">Choisissez un nom tel `=INCREASETIME` que plutôt `=INC`que.</span><span class="sxs-lookup"><span data-stu-id="b7180-130">Choose a name like `=INCREASETIME` rather than `=INC`.</span></span>
- <span data-ttu-id="b7180-131">Utilisez régulièrement les mêmes verbes pour les fonctions qui effectuent des actions similaires.</span><span class="sxs-lookup"><span data-stu-id="b7180-131">Consistently use the same verbs for functions which perform similar actions.</span></span> <span data-ttu-id="b7180-132">Par exemple, utilisez `=DELETEZIPCODE` and `=DELETEADDRESS`, et non `=DELETEZIPCODE` et `=REMOVEADDRESS`.</span><span class="sxs-lookup"><span data-stu-id="b7180-132">For example, use `=DELETEZIPCODE` and `=DELETEADDRESS`, rather than `=DELETEZIPCODE` and `=REMOVEADDRESS`.</span></span>
- <span data-ttu-id="b7180-133">Lorsque vous nommez une fonction de diffusion en continu, envisagez d’ajouter une note à cet effet `STREAM` dans la description de la fonction ou d’ajouter à la fin du nom de la fonction.</span><span class="sxs-lookup"><span data-stu-id="b7180-133">When naming a streaming function, consider adding a note to that effect in the description of the function or adding `STREAM` to the end of the function's name.</span></span>

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="localizing-function-names"></a><span data-ttu-id="b7180-134">Localisation des noms de fonction</span><span class="sxs-lookup"><span data-stu-id="b7180-134">Localizing function names</span></span>

<span data-ttu-id="b7180-135">Vous pouvez localiser vos noms de fonction pour différentes langues à l’aide de fichiers JSON distincts et remplacer les valeurs dans le fichier manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="b7180-135">You can localize your function names for different languages using separate JSON files and override values in your add-in's manifest file.</span></span> <span data-ttu-id="b7180-136">Nous vous recommandons de ne pas donner à vos fonctions `id` une `name` ou une fonction Excel intégrée dans un autre langage, car cela peut entraîner des conflits avec des fonctions localisées.</span><span class="sxs-lookup"><span data-stu-id="b7180-136">As a best practice, avoid giving your functions an `id` or `name` that is a built-in Excel function in another language as this could conflict with localized functions.</span></span>

<span data-ttu-id="b7180-137">Pour obtenir des informations complètes sur la localisation, voir [Localize Custom Functions](custom-functions-localize.md)</span><span class="sxs-lookup"><span data-stu-id="b7180-137">For full information on localizing, see [Localize custom functions](custom-functions-localize.md)</span></span>

## <a name="next-steps"></a><span data-ttu-id="b7180-138">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="b7180-138">Next steps</span></span>
<span data-ttu-id="b7180-139">Découvrez les [meilleures pratiques en matière de gestion des erreurs](custom-functions-errors.md).</span><span class="sxs-lookup"><span data-stu-id="b7180-139">Learn about [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="b7180-140">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b7180-140">See also</span></span>

* [<span data-ttu-id="b7180-141">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b7180-141">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="b7180-142">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="b7180-142">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="b7180-143">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="b7180-143">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
