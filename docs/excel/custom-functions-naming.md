---
ms.date: 02/08/2019
description: Découvrez les conditions requises pour les noms des fonctions personnalisées Excel et éviter les pièges de dénomination courants.
title: Instructions d'affectation de noms pour les fonctions personnalisées dans Excel (aperçu)
localization_priority: Normal
ms.openlocfilehash: bdf31879fb6e750fb9dea51f66c55dbc83a2dc90
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/22/2019
ms.locfileid: "30203846"
---
# <a name="naming-guidelines"></a><span data-ttu-id="b6d9a-103">Instructions d'affectation de noms</span><span class="sxs-lookup"><span data-stu-id="b6d9a-103">Naming guidelines</span></span>

<span data-ttu-id="b6d9a-104">Une fonction personnalisée est identifiée par un **ID** et une propriété de **nom** dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="b6d9a-104">A custom function is identified by an **id** and **name** property in the JSON metadata file.</span></span> <span data-ttu-id="b6d9a-105">L'ID de la fonction permet d'identifier de manière unique les fonctions personnalisées dans votre code JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b6d9a-105">The function id is used to uniquely identify custom functions in your JavaScript code.</span></span> <span data-ttu-id="b6d9a-106">Le nom de la fonction est utilisé comme nom complet qui apparaît pour un utilisateur dans Excel.</span><span class="sxs-lookup"><span data-stu-id="b6d9a-106">The function name is used as the display name that appears to a user in Excel.</span></span> <span data-ttu-id="b6d9a-107">Un nom de fonction peut différer de l'ID de fonction, par exemple à des fins de localisation.</span><span class="sxs-lookup"><span data-stu-id="b6d9a-107">A function name can differ from the function ID, such as for localization purposes.</span></span> <span data-ttu-id="b6d9a-108">Toutefois, en général, il doit rester identique à l'ID s'il n'y a aucune raison impérieuse qu'ils diffèrent.</span><span class="sxs-lookup"><span data-stu-id="b6d9a-108">But in general it should stay the same as the ID if there is no compelling reason for them to differ.</span></span>

<span data-ttu-id="b6d9a-109">Les noms de fonction et les ID de fonction partagent des exigences communes:</span><span class="sxs-lookup"><span data-stu-id="b6d9a-109">Function names and function IDs share some common requirements:</span></span>

- <span data-ttu-id="b6d9a-110">Elles doivent uniquement utiliser des caractères alphanumériques (y compris Unicode), les chiffres 0 à 9, des traits de soulignement et des points.</span><span class="sxs-lookup"><span data-stu-id="b6d9a-110">They must only use alphanumeric characters (including Unicode), the numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="b6d9a-111">Ils doivent commencer par une lettre et avoir une limite minimale de trois caractères.</span><span class="sxs-lookup"><span data-stu-id="b6d9a-111">They must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="b6d9a-112">Excel utilise des lettres majuscules pour les noms de fonctions intégrées ( `SUM`par exemple,).</span><span class="sxs-lookup"><span data-stu-id="b6d9a-112">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="b6d9a-113">Par conséquent, il est recommandé d'utiliser des lettres majuscules pour vos noms de fonction et ID de fonction personnalisés en tant que meilleure pratique.</span><span class="sxs-lookup"><span data-stu-id="b6d9a-113">Therefore, consider using uppercase letters for your custom function names and function IDs as a best practice.</span></span>

<span data-ttu-id="b6d9a-114">Les noms de fonction ne doivent pas porter le même nom que:</span><span class="sxs-lookup"><span data-stu-id="b6d9a-114">Function names shouldn't be named the same as:</span></span>

- <span data-ttu-id="b6d9a-115">Toutes les cellules comprises entre a1 et XFD1048576 ou toutes les cellules comprises entre R1C1 et R1048576C16384.</span><span class="sxs-lookup"><span data-stu-id="b6d9a-115">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="b6d9a-116">N'importe quelle fonction macro Excel 4,0 ( `RUN`telle `ECHO`que,).</span><span class="sxs-lookup"><span data-stu-id="b6d9a-116">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="b6d9a-117">Pour obtenir une liste complète de ces fonctions, consultez [cet article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).</span><span class="sxs-lookup"><span data-stu-id="b6d9a-117">For a full list of these functions, see [this article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="b6d9a-118">Conflits de noms</span><span class="sxs-lookup"><span data-stu-id="b6d9a-118">Naming conflicts</span></span>

<span data-ttu-id="b6d9a-119">Si le nom de votre fonction est identique à celui d'un nom de fonction dans un complément qui existe déjà, le **#REF!**</span><span class="sxs-lookup"><span data-stu-id="b6d9a-119">If your function name is the same as a function name in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="b6d9a-120">une erreur apparaît dans votre classeur.</span><span class="sxs-lookup"><span data-stu-id="b6d9a-120">error will appear in your workbook.</span></span>

<span data-ttu-id="b6d9a-121">Pour résoudre un conflit de nom, modifiez le nom dans votre complément et réessayez.</span><span class="sxs-lookup"><span data-stu-id="b6d9a-121">To fix a name conflict, change the name in your add-in and try the function again.</span></span> <span data-ttu-id="b6d9a-122">Vous pouvez également désinstaller le complément avec le nom conflictuel.</span><span class="sxs-lookup"><span data-stu-id="b6d9a-122">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="b6d9a-123">Ou, si vous testez votre complément dans différents environnements, essayez d'utiliser un espace de noms différent pour différencier votre fonction (par exemple, NAMESPACE_NAMEOFFUNCTION).</span><span class="sxs-lookup"><span data-stu-id="b6d9a-123">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as NAMESPACE_NAMEOFFUNCTION).</span></span>

<span data-ttu-id="b6d9a-124">Réfléchissez également à la façon dont vous souhaitez que les personnes utilisent les fonctions dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="b6d9a-124">Also consider how you'd like people to use the functions within your add-in.</span></span> <span data-ttu-id="b6d9a-125">Dans de nombreux cas, il est logique d'ajouter plusieurs arguments à une fonction plutôt que de créer plusieurs fonctions avec des noms identiques ou similaires.</span><span class="sxs-lookup"><span data-stu-id="b6d9a-125">In many cases, it makes sense to add multiple arguments to a function rather than create multiple functions with the same or similar names.</span></span>

## <a name="see-also"></a><span data-ttu-id="b6d9a-126">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b6d9a-126">See also</span></span>

* [<span data-ttu-id="b6d9a-127">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b6d9a-127">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="b6d9a-128">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b6d9a-128">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="b6d9a-129">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="b6d9a-129">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="b6d9a-130">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="b6d9a-130">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
