---
ms.date: 01/08/2020
description: Créez une fonction personnalisée Excel pour votre Complément Office.
title: Créer des fonctions personnalisées dans Excel
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 97037f201a237cdc6dae551552a0a1609a58b34c
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839872"
---
# <a name="create-custom-functions-in-excel"></a><span data-ttu-id="65210-103">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="65210-103">Create custom functions in Excel</span></span>

<span data-ttu-id="65210-104">Les fonctions personnalisées permettent aux développeurs d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="65210-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="65210-105">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="65210-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="65210-106">L’image animée suivante montre votre classeur appelant une fonction que vous avez créée avec JavaScript ou Typescript.</span><span class="sxs-lookup"><span data-stu-id="65210-106">The following animated image shows your workbook calling a function you've created with JavaScript or Typescript.</span></span> <span data-ttu-id="65210-107">Dans cet exemple, la fonction personnalisée `=MYFUNCTION.SPHEREVOLUME` calcule le volume d’une sphère.</span><span class="sxs-lookup"><span data-stu-id="65210-107">In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

<span data-ttu-id="65210-108">Le code suivant définit la fonction personnalisée `=MYFUNCTION.SPHEREVOLUME`.</span><span class="sxs-lookup"><span data-stu-id="65210-108">The following code defines the custom function `=MYFUNCTION.SPHEREVOLUME`.</span></span>

```js
/**
 * Returns the volume of a sphere.
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return Math.pow(radius, 3) * 4 * Math.PI / 3;
}
```

> [!TIP]
> <span data-ttu-id="65210-109">Si votre complément de fonction personnalisée utilise un volet Office ou un bouton du ruban, outre l’exécution du code de fonction personnalisée, vous devez configurer un runtime JavaScript partagé.</span><span class="sxs-lookup"><span data-stu-id="65210-109">If your custom function add-in will use a task pane or a ribbon button, in addition to running custom function code, you will need to set up a shared JavaScript runtime.</span></span> <span data-ttu-id="65210-110">Pour plus d’information, consultez [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="65210-110">See [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md) to learn more.</span></span>

## <a name="how-a-custom-function-is-defined-in-code"></a><span data-ttu-id="65210-111">Comment une fonction personnalisée est définie dans le code</span><span class="sxs-lookup"><span data-stu-id="65210-111">How a custom function is defined in code</span></span>

<span data-ttu-id="65210-112">Si vous utilisez le [générateur de Yo Office](https://github.com/OfficeDev/generator-office) pour créer un projet de complément de fonctions personnalisées Excel, il crée des fichiers qui contrôlent totalement vos fonctions, et volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="65210-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, it creates files which control your functions and task pane.</span></span> <span data-ttu-id="65210-113">Nous allons vous concentrer sur les fichiers importants pour les fonctions personnalisées :</span><span class="sxs-lookup"><span data-stu-id="65210-113">We'll concentrate on the files that are important to custom functions:</span></span>

| <span data-ttu-id="65210-114">File</span><span class="sxs-lookup"><span data-stu-id="65210-114">File</span></span> | <span data-ttu-id="65210-115">Format de fichier</span><span class="sxs-lookup"><span data-stu-id="65210-115">File format</span></span> | <span data-ttu-id="65210-116">Description</span><span class="sxs-lookup"><span data-stu-id="65210-116">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="65210-117">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="65210-117">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="65210-118">ou</span><span class="sxs-lookup"><span data-stu-id="65210-118">or</span></span><br/><span data-ttu-id="65210-119">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="65210-119">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="65210-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="65210-120">JavaScript</span></span><br/><span data-ttu-id="65210-121">ou</span><span class="sxs-lookup"><span data-stu-id="65210-121">or</span></span><br/><span data-ttu-id="65210-122">TypeScript</span><span class="sxs-lookup"><span data-stu-id="65210-122">TypeScript</span></span> | <span data-ttu-id="65210-123">Contient le code qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="65210-123">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="65210-124">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="65210-124">**./src/functions/functions.html**</span></span> | <span data-ttu-id="65210-125">HTML</span><span class="sxs-lookup"><span data-stu-id="65210-125">HTML</span></span> | <span data-ttu-id="65210-126">Fournit une référence&lt;script&gt; au fichier JavaScript qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="65210-126">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="65210-127">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="65210-127">**./manifest.xml**</span></span> | <span data-ttu-id="65210-128">XML</span><span class="sxs-lookup"><span data-stu-id="65210-128">XML</span></span> | <span data-ttu-id="65210-129">Indique l’emplacement de plusieurs fichiers utilisés par votre fonction personnalisée, tels que les fonctions personnalisées JavaScript, JSON et HTML.</span><span class="sxs-lookup"><span data-stu-id="65210-129">Specifies the location of multiple files that your custom function use, such as the custom functions JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="65210-130">Il répertorie également les emplacements des fichiers du volet Office, des fichiers de commandes et indique le runtime que vos fonctions personnalisées doivent utiliser.</span><span class="sxs-lookup"><span data-stu-id="65210-130">It also lists the locations of task pane files, command files, and specifies which runtime your custom functions should use.</span></span> |

### <a name="script-file"></a><span data-ttu-id="65210-131">Fichier de script</span><span class="sxs-lookup"><span data-stu-id="65210-131">Script file</span></span>

<span data-ttu-id="65210-132">Le fichier de script (**./src/functions/functions.js** ou **./src/functions/functions.ts**) contient le code qui définit des fonctions personnalisées et des commentaires qui définissent la fonction.</span><span class="sxs-lookup"><span data-stu-id="65210-132">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts**) contains the code that defines custom functions and comments which define the function.</span></span>

<span data-ttu-id="65210-133">Le code suivant définit la fonction personnalisée `add`.</span><span class="sxs-lookup"><span data-stu-id="65210-133">The following code defines the custom function `add`.</span></span> <span data-ttu-id="65210-134">Les commentaires du code sont utilisés pour générer un fichier de métadonnées JSON décrivant la fonction personnalisée pour Excel.</span><span class="sxs-lookup"><span data-stu-id="65210-134">The code comments are used to generate a JSON metadata file that describes the custom function to Excel.</span></span> <span data-ttu-id="65210-135">Le commentaire obligatoire `@customfunction` est déclaré en premier, pour indiquer qu’il s’agit d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="65210-135">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="65210-136">Deux paramètres sont ensuite déclarés, `first` et `second`, suivis de leurs propriétés de `description` .</span><span class="sxs-lookup"><span data-stu-id="65210-136">Next, two parameters are declared, `first` and `second`, followed by their `description` properties.</span></span> <span data-ttu-id="65210-137">Enfin, une description `returns` est fournie.</span><span class="sxs-lookup"><span data-stu-id="65210-137">Finally, a `returns` description is given.</span></span> <span data-ttu-id="65210-138">Pour plus d’informations sur les commentaires requis pour votre fonction personnalisée, voir [Générer automatiquement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="65210-138">For more information about what comments are required for your custom function, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number.
 * @param second Second number.
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}
```

### <a name="manifest-file"></a><span data-ttu-id="65210-139">Fichier manifeste</span><span class="sxs-lookup"><span data-stu-id="65210-139">Manifest file</span></span>

<span data-ttu-id="65210-140">Le fichier manifeste XML pour un complément qui définit des fonctions personnalisées (**./manifest.xml** dans le projet que le générateur de bureau Yo crée) effectue plusieurs opérations :</span><span class="sxs-lookup"><span data-stu-id="65210-140">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) does several things:</span></span>

- <span data-ttu-id="65210-141">Définit l’espace de noms pour vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="65210-141">Defines the namespace for your custom functions.</span></span> <span data-ttu-id="65210-142">Un espace de noms s’ajoute à vos fonctions personnalisées pour aider les clients à identifier vos fonctions dans le cadre de votre complément.</span><span class="sxs-lookup"><span data-stu-id="65210-142">A namespace prepends itself to your custom functions to help customers identify your functions as part of your add-in.</span></span>
- <span data-ttu-id="65210-143">Utilise les éléments `<ExtensionPoint>` et `<Resources>` qui sont propres à un manifeste de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="65210-143">Uses `<ExtensionPoint>` and `<Resources>` elements that are unique to a custom functions manifest.</span></span> <span data-ttu-id="65210-144">Ces éléments contiennent les informations relatives aux emplacements des fichiers JavaScript, JSON et HTML.</span><span class="sxs-lookup"><span data-stu-id="65210-144">These elements contain the information about the locations of the JavaScript, JSON, and HTML files.</span></span>
- <span data-ttu-id="65210-145">Spécifie le runtime à utiliser pour votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="65210-145">Specifies which runtime to use for your custom function.</span></span> <span data-ttu-id="65210-146">Nous vous recommandons de toujours utiliser une exécution partagée, sauf si vous avez un besoin spécifique d’autre runtime, car un runtime partagé autorise le partage de données entre les fonctions et le volet Office.</span><span class="sxs-lookup"><span data-stu-id="65210-146">We recommend always using a shared runtime unless you have a specific need for another runtime, because a shared runtime allows for the sharing of data between functions and the task pane.</span></span> <span data-ttu-id="65210-147">Notez que l’utilisation d’un runtime partagé signifie que votre complément utilise Internet Explorer 11, et non Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="65210-147">Note that using a shared runtime means your add-in will use Internet Explorer 11, not Microsoft Edge.</span></span>

<span data-ttu-id="65210-148">Si vous utilisez le générateur Yo Office pour créer des fichiers, nous vous recommandons d’ajuster votre manifeste pour utiliser un runtime partagé, car il ne s’agit pas de la valeur par défaut pour ces fichiers.</span><span class="sxs-lookup"><span data-stu-id="65210-148">If you are using the Yo Office generator to create files, we recommend adjusting your manifest to use a shared runtime, as this is not the default for these files.</span></span> <span data-ttu-id="65210-149">Pour modifier votre manifeste, suivez les instructions dans [Configurer votre complément Excel pour utiliser un runtime JavaScript partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="65210-149">To change your manifest, follow the instructions in [Configure your Excel add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="65210-150">Pour afficher un manifeste de travail complet à partir d’un exemple de complément, consultez [ce référentiel GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="65210-150">To see a full working manifest from a sample add-in, see [this Github repository](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).</span></span>

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a><span data-ttu-id="65210-151">Co-édition</span><span class="sxs-lookup"><span data-stu-id="65210-151">Coauthoring</span></span>

<span data-ttu-id="65210-152">Excel sur le web et Windows connecté à un abonnement Microsoft 365 vous permettent de co-éditer dans Excel.</span><span class="sxs-lookup"><span data-stu-id="65210-152">Excel on the web and Windows connected to a Microsoft 365 subscription allow you to coauthor in Excel.</span></span> <span data-ttu-id="65210-153">Si votre classeur utilise une fonction personnalisée, votre collègue coauteur est invité à charger le complément de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="65210-153">If your workbook uses a custom function, your coauthoring colleague is prompted to load the custom function's add-in.</span></span> <span data-ttu-id="65210-154">Quand vous avez tous les deux chargé le complément, la fonction personnalisée partage les résultats via la co-édition.</span><span class="sxs-lookup"><span data-stu-id="65210-154">Once you both have loaded the add-in, the custom function shares results through coauthoring.</span></span>

<span data-ttu-id="65210-155">Pour plus d’informations sur la co-création, voir [À propos de la co-création dans Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span><span class="sxs-lookup"><span data-stu-id="65210-155">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="next-steps"></a><span data-ttu-id="65210-156">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="65210-156">Next steps</span></span>

<span data-ttu-id="65210-157">Vous voulez essayer les fonctions personnalisées ?</span><span class="sxs-lookup"><span data-stu-id="65210-157">Want to try out custom functions?</span></span> <span data-ttu-id="65210-158">Consultez la documentation sur le [démarrage rapide de fonction personnalisée](../quickstarts/excel-custom-functions-quickstart.md) ou le [didacticiel sur les fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="65210-158">Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.</span></span>

<span data-ttu-id="65210-159">Un autre moyen simple d’essayer des fonctions personnalisées consiste à utiliser [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), un complément qui vous permet d’expérimenter des fonctions personnalisées directement dans Excel.</span><span class="sxs-lookup"><span data-stu-id="65210-159">Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel.</span></span> <span data-ttu-id="65210-160">Vous pouvez essayer de créer votre propre fonction personnalisée ou utiliser les exemples fournis.</span><span class="sxs-lookup"><span data-stu-id="65210-160">You can try out creating your own custom function or play with the provided samples.</span></span>

## <a name="see-also"></a><span data-ttu-id="65210-161">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="65210-161">See also</span></span> 
* [<span data-ttu-id="65210-162">Découvrez le programme pour les développeurs Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="65210-162">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
* [<span data-ttu-id="65210-163">Ensembles de besoins de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="65210-163">Custom functions requirement sets</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="65210-164">Règles de noms des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="65210-164">Custom functions naming guidelines</span></span>](custom-functions-naming.md)
* [<span data-ttu-id="65210-165">Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="65210-165">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
* [<span data-ttu-id="65210-166">Configurer votre complément Office pour utiliser un runtime JavaScript partagé</span><span class="sxs-lookup"><span data-stu-id="65210-166">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
