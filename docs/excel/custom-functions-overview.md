---
ms.date: 05/17/2020
description: Créez une fonction personnalisée Excel pour votre Complément Office
title: Créer des fonctions personnalisées dans Excel
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 4f8416b9058def9dcb4998fb2f31684b59276ac4
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609281"
---
# <a name="create-custom-functions-in-excel"></a><span data-ttu-id="8e9b0-103">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="8e9b0-103">Create custom functions in Excel</span></span>

<span data-ttu-id="8e9b0-104">Les fonctions personnalisées permettent aux développeurs d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="8e9b0-105">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="8e9b0-106">L’image animée suivante montre votre classeur appelant une fonction que vous avez créée avec JavaScript ou Typescript.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-106">The following animated image shows your workbook calling a function you've created with JavaScript or Typescript.</span></span> <span data-ttu-id="8e9b0-107">Dans cet exemple, la fonction personnalisée `=MYFUNCTION.SPHEREVOLUME` calcule le volume d’une sphère.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-107">In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

<span data-ttu-id="8e9b0-108">Le code suivant définit la fonction personnalisée `=MYFUNCTION.SPHEREVOLUME`.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-108">The following code defines the custom function `=MYFUNCTION.SPHEREVOLUME`.</span></span>

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

> [!NOTE]
> <span data-ttu-id="8e9b0-109">La section [problèmes connus](#known-issues)plus loin dans cet article indique les limitations en cours de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-109">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="how-a-custom-function-is-defined-in-code"></a><span data-ttu-id="8e9b0-110">Comment une fonction personnalisée est définie dans le code</span><span class="sxs-lookup"><span data-stu-id="8e9b0-110">How a custom function is defined in code</span></span>

<span data-ttu-id="8e9b0-111">Si vous utilisez le [Générateur Yo Office](https://github.com/OfficeDev/generator-office) pour créer un projet de complément de fonctions personnalisées Excel, il crée des fichiers qui contrôlent vos fonctions et volet de tâches.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-111">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, it creates files which control your functions and task pane.</span></span> <span data-ttu-id="8e9b0-112">Nous allons vous concentrer sur les fichiers importants pour les fonctions personnalisées :</span><span class="sxs-lookup"><span data-stu-id="8e9b0-112">We'll concentrate on the files that are important to custom functions:</span></span>

| <span data-ttu-id="8e9b0-113">File</span><span class="sxs-lookup"><span data-stu-id="8e9b0-113">File</span></span> | <span data-ttu-id="8e9b0-114">Format de fichier</span><span class="sxs-lookup"><span data-stu-id="8e9b0-114">File format</span></span> | <span data-ttu-id="8e9b0-115">Description</span><span class="sxs-lookup"><span data-stu-id="8e9b0-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="8e9b0-116">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="8e9b0-116">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="8e9b0-117">ou</span><span class="sxs-lookup"><span data-stu-id="8e9b0-117">or</span></span><br/><span data-ttu-id="8e9b0-118">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="8e9b0-118">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="8e9b0-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="8e9b0-119">JavaScript</span></span><br/><span data-ttu-id="8e9b0-120">ou</span><span class="sxs-lookup"><span data-stu-id="8e9b0-120">or</span></span><br/><span data-ttu-id="8e9b0-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="8e9b0-121">TypeScript</span></span> | <span data-ttu-id="8e9b0-122">Contient le code qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="8e9b0-123">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="8e9b0-123">**./src/functions/functions.html**</span></span> | <span data-ttu-id="8e9b0-124">HTML</span><span class="sxs-lookup"><span data-stu-id="8e9b0-124">HTML</span></span> | <span data-ttu-id="8e9b0-125">Fournit une référence&lt;script&gt; au fichier JavaScript qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-125">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="8e9b0-126">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="8e9b0-126">**./manifest.xml**</span></span> | <span data-ttu-id="8e9b0-127">XML</span><span class="sxs-lookup"><span data-stu-id="8e9b0-127">XML</span></span> | <span data-ttu-id="8e9b0-128">Spécifie l’emplacement de plusieurs fichiers que votre fonction personnalisée utilise, tels que les fichiers JavaScript, JSON et HTML des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-128">Specifies the location of multiple files that your custom function use, such as the custom functions JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="8e9b0-129">Il répertorie également les emplacements des fichiers de volet de tâches, des fichiers de commandes et spécifie le runtime que vos fonctions personnalisées doivent utiliser.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-129">It also lists the locations of task pane files, command files, and specifies which runtime your custom functions should use.</span></span> |

### <a name="script-file"></a><span data-ttu-id="8e9b0-130">Fichier de script</span><span class="sxs-lookup"><span data-stu-id="8e9b0-130">Script file</span></span>

<span data-ttu-id="8e9b0-131">Le fichier de script (**./src/functions/functions.js** ou **./src/functions/functions.ts**) contient le code qui définit des fonctions personnalisées et des commentaires qui définissent la fonction.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-131">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts**) contains the code that defines custom functions and comments which define the function.</span></span>

<span data-ttu-id="8e9b0-132">Le code suivant définit la fonction personnalisée `add`.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-132">The following code defines the custom function `add`.</span></span> <span data-ttu-id="8e9b0-133">Les commentaires du code sont utilisés pour générer un fichier de métadonnées JSON décrivant la fonction personnalisée pour Excel.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-133">The code comments are used to generate a JSON metadata file that describes the custom function to Excel.</span></span> <span data-ttu-id="8e9b0-134">Le commentaire obligatoire `@customfunction` est déclaré en premier, pour indiquer qu’il s’agit d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-134">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="8e9b0-135">Ensuite, deux paramètres sont déclarés `first` et `second` , suivis de leurs `description` Propriétés.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-135">Next, two parameters are declared, `first` and `second`, followed by their `description` properties.</span></span> <span data-ttu-id="8e9b0-136">Enfin, une description `returns` est fournie.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-136">Finally, a `returns` description is given.</span></span> <span data-ttu-id="8e9b0-137">Pour plus d’informations sur les commentaires requis pour votre fonction personnalisée, voir [Créer des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="8e9b0-137">For more information about what comments are required for your custom function, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

### <a name="manifest-file"></a><span data-ttu-id="8e9b0-138">Fichier manifeste</span><span class="sxs-lookup"><span data-stu-id="8e9b0-138">Manifest file</span></span>

<span data-ttu-id="8e9b0-139">Le fichier manifeste XML d’un complément qui définit des fonctions personnalisées (**./manifest.xml** dans le projet créé par le générateur Yo Office) effectue plusieurs actions :</span><span class="sxs-lookup"><span data-stu-id="8e9b0-139">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) does several things:</span></span>

- <span data-ttu-id="8e9b0-140">Définit l’espace de noms pour vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-140">Defines the namespace for your custom functions.</span></span> <span data-ttu-id="8e9b0-141">Un espace de noms s’ajoute à vos fonctions personnalisées pour aider les clients à identifier vos fonctions dans le cadre de votre complément.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-141">A namespace prepends itself to your custom functions to help customers identify your functions as part of your add-in.</span></span>
- <span data-ttu-id="8e9b0-142">Utilisations `<ExtensionPoint>` et `<Resources>` éléments propres à un manifeste de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-142">Uses `<ExtensionPoint>` and `<Resources>` elements that are unique to a custom functions manifest.</span></span> <span data-ttu-id="8e9b0-143">Ces éléments contiennent des informations sur les emplacements des fichiers JavaScript, JSON et HTML.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-143">These elements contain the information about the locations of the JavaScript, JSON, and HTML files.</span></span>
- <span data-ttu-id="8e9b0-144">Spécifie le runtime à utiliser pour votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-144">Specifies which runtime to use for your custom function.</span></span> <span data-ttu-id="8e9b0-145">Nous vous recommandons de toujours utiliser un runtime partagé, sauf si vous avez besoin d’un autre Runtime spécifique, car un runtime partagé autorise le partage des données entre les fonctions et le volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-145">We recommend always using a shared runtime unless you have a specific need for another runtime, because a shared runtime allows for the sharing of data between functions and the task pane.</span></span>

<span data-ttu-id="8e9b0-146">Si vous utilisez le générateur Yo Office pour créer des fichiers, nous vous recommandons d’ajuster votre manifeste afin qu’il utilise un runtime partagé, car il ne s’agit pas de la valeur par défaut pour ces fichiers.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-146">If you are using the Yo Office generator to create files, we recommend adjusting your manifest to use a shared runtime, as this is not the default for these files.</span></span> <span data-ttu-id="8e9b0-147">Pour modifier votre manifeste, suivez les instructions de la procédure de [configuration de votre complément Excel pour utiliser un Runtime JavaScript partagé](./configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="8e9b0-147">To change your manifest, follow the instructions in [Configure your Excel add-in to use a shared JavaScript runtime](./configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="8e9b0-148">Pour afficher un manifeste de travail complet à partir d’un exemple de complément, reportez-vous à [ce référentiel GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="8e9b0-148">To see a full working manifest from a sample add-in, see [this Github repository](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).</span></span>

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a><span data-ttu-id="8e9b0-149">Co-création</span><span class="sxs-lookup"><span data-stu-id="8e9b0-149">Coauthoring</span></span>

<span data-ttu-id="8e9b0-150">Excel sur le Web et Windows connecté à un abonnement Office 365 vous permettent de co-auteur dans Excel.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-150">Excel on the web and Windows connected to an Office 365 subscription allow you to coauthor in Excel.</span></span> <span data-ttu-id="8e9b0-151">Si votre classeur utilise une fonction personnalisée, votre collègue de co-création est invité à charger le complément de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-151">If your workbook uses a custom function, your coauthoring colleague is prompted to load the custom function's add-in.</span></span> <span data-ttu-id="8e9b0-152">Une fois que vous avez chargé le complément, la fonction personnalisée partage les résultats par le biais de la co-création.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-152">Once you both have loaded the add-in, the custom function shares results through coauthoring.</span></span>

<span data-ttu-id="8e9b0-153">Pour plus d’informations sur la co-création, voir [À propos de la co-création dans Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span><span class="sxs-lookup"><span data-stu-id="8e9b0-153">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="known-issues"></a><span data-ttu-id="8e9b0-154">Problèmes connus</span><span class="sxs-lookup"><span data-stu-id="8e9b0-154">Known issues</span></span>

<span data-ttu-id="8e9b0-155">Consulter les problèmes connus sur notre[repo GitHub Fonctions Excel Personnalisées](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span><span class="sxs-lookup"><span data-stu-id="8e9b0-155">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="8e9b0-156">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="8e9b0-156">Next steps</span></span>

<span data-ttu-id="8e9b0-157">Vous voulez essayer les fonctions personnalisées ?</span><span class="sxs-lookup"><span data-stu-id="8e9b0-157">Want to try out custom functions?</span></span> <span data-ttu-id="8e9b0-158">Consultez la documentation sur le [démarrage rapide de fonction personnalisée](../quickstarts/excel-custom-functions-quickstart.md) ou le [didacticiel sur les fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="8e9b0-158">Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.</span></span>

<span data-ttu-id="8e9b0-159">Un autre moyen simple d’essayer des fonctions personnalisées consiste à utiliser [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), un complément qui vous permet d’expérimenter des fonctions personnalisées directement dans Excel.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-159">Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel.</span></span> <span data-ttu-id="8e9b0-160">Vous pouvez essayer de créer votre propre fonction personnalisée ou utiliser les exemples fournis.</span><span class="sxs-lookup"><span data-stu-id="8e9b0-160">You can try out creating your own custom function or play with the provided samples.</span></span>

## <a name="see-also"></a><span data-ttu-id="8e9b0-161">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8e9b0-161">See also</span></span> 
* [<span data-ttu-id="8e9b0-162">Configuration requise de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="8e9b0-162">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="8e9b0-163">Instructions d’attribution de noms</span><span class="sxs-lookup"><span data-stu-id="8e9b0-163">Naming guidelines</span></span>](custom-functions-naming.md)
* [<span data-ttu-id="8e9b0-164">Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="8e9b0-164">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
