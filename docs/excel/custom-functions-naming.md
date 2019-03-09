---
ms.date: 02/08/2019
description: Découvrez les conditions requises pour les noms des fonctions personnalisées Excel et éviter les pièges de dénomination courants.
title: Instructions d'affectation de noms pour les fonctions personnalisées dans Excel (aperçu)
localization_priority: Normal
ms.openlocfilehash: 954753c35d2df59093661e3b8e92adfa1302e595
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512838"
---
# <a name="naming-guidelines"></a><span data-ttu-id="49999-103">Instructions d'affectation de noms</span><span class="sxs-lookup"><span data-stu-id="49999-103">Naming guidelines</span></span>

<span data-ttu-id="49999-104">Une fonction personnalisée est identifiée par un **ID** et une propriété de **nom** dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="49999-104">A custom function is identified by an **id** and **name** property in the JSON metadata file.</span></span> <span data-ttu-id="49999-105">L'ID de la fonction permet d'identifier de manière unique les fonctions personnalisées dans votre code JavaScript.</span><span class="sxs-lookup"><span data-stu-id="49999-105">The function id is used to uniquely identify custom functions in your JavaScript code.</span></span> <span data-ttu-id="49999-106">Le nom de la fonction est utilisé comme nom complet qui apparaît pour un utilisateur dans Excel.</span><span class="sxs-lookup"><span data-stu-id="49999-106">The function name is used as the display name that appears to a user in Excel.</span></span> <span data-ttu-id="49999-107">Un nom de fonction peut différer de l'ID de fonction, par exemple à des fins de localisation.</span><span class="sxs-lookup"><span data-stu-id="49999-107">A function name can differ from the function ID, such as for localization purposes.</span></span> <span data-ttu-id="49999-108">Toutefois, en général, il doit rester identique à l'ID s'il n'y a aucune raison impérieuse qu'ils diffèrent.</span><span class="sxs-lookup"><span data-stu-id="49999-108">But in general it should stay the same as the ID if there is no compelling reason for them to differ.</span></span>

<span data-ttu-id="49999-109">Les noms de fonction et les ID de fonction partagent des exigences communes:</span><span class="sxs-lookup"><span data-stu-id="49999-109">Function names and function IDs share some common requirements:</span></span>

- <span data-ttu-id="49999-110">Les ID de fonction ne peuvent utiliser que les caractères A à Z, les chiffres 0 à 9, les traits de soulignement et les points.</span><span class="sxs-lookup"><span data-stu-id="49999-110">Function ids may only use characters A through Z, numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="49999-111">Les noms de fonction peuvent utiliser n'importe quel caractère alphabétique Unicode, des traits de soulignement et des points.</span><span class="sxs-lookup"><span data-stu-id="49999-111">Function names may use any Unicode alphabetic characters, underscores, and periods.</span></span>

- <span data-ttu-id="49999-112">Ils doivent commencer par une lettre et avoir une limite minimale de trois caractères.</span><span class="sxs-lookup"><span data-stu-id="49999-112">They must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="49999-113">Excel utilise des lettres majuscules pour les noms de fonctions intégrées ( `SUM`par exemple,).</span><span class="sxs-lookup"><span data-stu-id="49999-113">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="49999-114">Par conséquent, il est recommandé d'utiliser des lettres majuscules pour vos noms de fonction et ID de fonction personnalisés en tant que meilleure pratique.</span><span class="sxs-lookup"><span data-stu-id="49999-114">Therefore, consider using uppercase letters for your custom function names and function IDs as a best practice.</span></span>

<span data-ttu-id="49999-115">Les noms de fonction ne doivent pas porter le même nom que:</span><span class="sxs-lookup"><span data-stu-id="49999-115">Function names shouldn't be named the same as:</span></span>

- <span data-ttu-id="49999-116">Toutes les cellules comprises entre a1 et XFD1048576 ou toutes les cellules comprises entre R1C1 et R1048576C16384.</span><span class="sxs-lookup"><span data-stu-id="49999-116">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="49999-117">N'importe quelle fonction macro Excel 4,0 ( `RUN`telle `ECHO`que,).</span><span class="sxs-lookup"><span data-stu-id="49999-117">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="49999-118">Pour obtenir une liste complète de ces fonctions, consultez [cet article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).</span><span class="sxs-lookup"><span data-stu-id="49999-118">For a full list of these functions, see [this article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="49999-119">Conflits de noms</span><span class="sxs-lookup"><span data-stu-id="49999-119">Naming conflicts</span></span>

<span data-ttu-id="49999-120">Si le nom de votre fonction est identique à celui d'un nom de fonction dans un complément qui existe déjà, le **#REF!**</span><span class="sxs-lookup"><span data-stu-id="49999-120">If your function name is the same as a function name in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="49999-121">une erreur apparaît dans votre classeur.</span><span class="sxs-lookup"><span data-stu-id="49999-121">error will appear in your workbook.</span></span>

<span data-ttu-id="49999-122">Pour résoudre un conflit de nom, modifiez le nom dans votre complément et réessayez.</span><span class="sxs-lookup"><span data-stu-id="49999-122">To fix a name conflict, change the name in your add-in and try the function again.</span></span> <span data-ttu-id="49999-123">Vous pouvez également désinstaller le complément avec le nom conflictuel.</span><span class="sxs-lookup"><span data-stu-id="49999-123">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="49999-124">Ou, si vous testez votre complément dans différents environnements, essayez d'utiliser un espace de noms différent pour différencier votre fonction (par exemple, NAMESPACE_NAMEOFFUNCTION).</span><span class="sxs-lookup"><span data-stu-id="49999-124">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as NAMESPACE_NAMEOFFUNCTION).</span></span>

<span data-ttu-id="49999-125">Réfléchissez également à la façon dont vous souhaitez que les personnes utilisent les fonctions dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="49999-125">Also consider how you'd like people to use the functions within your add-in.</span></span> <span data-ttu-id="49999-126">Dans de nombreux cas, il est logique d'ajouter plusieurs arguments à une fonction plutôt que de créer plusieurs fonctions avec des noms identiques ou similaires.</span><span class="sxs-lookup"><span data-stu-id="49999-126">In many cases, it makes sense to add multiple arguments to a function rather than create multiple functions with the same or similar names.</span></span>

## <a name="see-also"></a><span data-ttu-id="49999-127">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="49999-127">See also</span></span>

* [<span data-ttu-id="49999-128">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="49999-128">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="49999-129">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="49999-129">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="49999-130">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="49999-130">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="49999-131">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="49999-131">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
