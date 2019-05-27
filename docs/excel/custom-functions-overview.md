---
ms.date: 05/15/2019
description: Créer des fonctions personnalisées dans Excel à l’aide de JavaScript.
title: Créer des fonctions personnalisées dans Excel
localization_priority: Priority
ms.openlocfilehash: 3eeedd482a432166a7fa26eff6da4b075847a292
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432168"
---
# <a name="create-custom-functions-in-excel"></a><span data-ttu-id="642a2-103">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="642a2-103">Create custom functions in Excel</span></span> 

<span data-ttu-id="642a2-104">Les fonctions personnalisées permettent aux développeurs d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="642a2-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="642a2-105">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="642a2-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="642a2-106">Cet article explique comment créer des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="642a2-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="642a2-107">L’image animée suivante montre votre classeur appelant une fonction que vous avez créée avec JavaScript ou Typescript.</span><span class="sxs-lookup"><span data-stu-id="642a2-107">The following animated image shows your workbook calling a function you've created with JavaScript or Typescript.</span></span> <span data-ttu-id="642a2-108">Dans cet exemple, la fonction personnalisée `=MYFUNCTION.SPHEREVOLUME` calcule le volume d’une sphère.</span><span class="sxs-lookup"><span data-stu-id="642a2-108">In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

<span data-ttu-id="642a2-109">Le code suivant définit la fonction personnalisée `=MYFUNCTION.SPHEREVOLUME`.</span><span class="sxs-lookup"><span data-stu-id="642a2-109">The following code defines the custom function `=MYFUNCTION.SPHEREVOLUME`.</span></span>

```js
/**
 * Returns the volume of a sphere. 
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return Math.pow(radius, 3) * 4 * Math.PI / 3;
}
CustomFunctions.associate("SPHEREVOLUME", sphereVolume)
```

> [!NOTE]
> <span data-ttu-id="642a2-110">La section [problèmes connus](#known-issues)plus loin dans cet article indique les limitations en cours de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="642a2-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="how-a-custom-function-is-defined-in-code"></a><span data-ttu-id="642a2-111">Comment une fonction personnalisée est définie dans le code</span><span class="sxs-lookup"><span data-stu-id="642a2-111">How a custom function is defined in code</span></span>

<span data-ttu-id="642a2-112">Si vous utilisez le [générateur de Yo Office](https://github.com/OfficeDev/generator-office) pour créer un projet de complément de fonctions personnalisées Excel, vous constaterez qu’il crée des fichiers qui contrôlent totalement vos fonctions, votre volet des tâches et votre complément.</span><span class="sxs-lookup"><span data-stu-id="642a2-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll find that it creates files which control your functions, your task pane, and your add-in overall.</span></span> <span data-ttu-id="642a2-113">Nous allons vous concentrer sur les fichiers importants pour les fonctions personnalisées :</span><span class="sxs-lookup"><span data-stu-id="642a2-113">We'll concentrate on the files that are important to custom functions:</span></span>

| <span data-ttu-id="642a2-114">File</span><span class="sxs-lookup"><span data-stu-id="642a2-114">File</span></span> | <span data-ttu-id="642a2-115">Format de fichier</span><span class="sxs-lookup"><span data-stu-id="642a2-115">File format</span></span> | <span data-ttu-id="642a2-116">Description</span><span class="sxs-lookup"><span data-stu-id="642a2-116">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="642a2-117">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="642a2-117">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="642a2-118">ou</span><span class="sxs-lookup"><span data-stu-id="642a2-118">or</span></span><br/><span data-ttu-id="642a2-119">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="642a2-119">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="642a2-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="642a2-120">JavaScript</span></span><br/><span data-ttu-id="642a2-121">ou</span><span class="sxs-lookup"><span data-stu-id="642a2-121">or</span></span><br/><span data-ttu-id="642a2-122">TypeScript</span><span class="sxs-lookup"><span data-stu-id="642a2-122">TypeScript</span></span> | <span data-ttu-id="642a2-123">Contient le code qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="642a2-123">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="642a2-124">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="642a2-124">**./src/functions/functions.html**</span></span> | <span data-ttu-id="642a2-125">HTML</span><span class="sxs-lookup"><span data-stu-id="642a2-125">HTML</span></span> | <span data-ttu-id="642a2-126">Fournit une référence&lt;script&gt; au fichier JavaScript qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="642a2-126">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="642a2-127">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="642a2-127">**./manifest.xml**</span></span> | <span data-ttu-id="642a2-128">XML</span><span class="sxs-lookup"><span data-stu-id="642a2-128">XML</span></span> | <span data-ttu-id="642a2-129">Spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers JavaScript et HTML qui figurent plus haut dans ce tableau.</span><span class="sxs-lookup"><span data-stu-id="642a2-129">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript and HTML files that are listed previously in this table.</span></span> <span data-ttu-id="642a2-130">Répertorie également les emplacements des autres fichiers que votre complément pourrait utiliser, tels que les fichiers du volet des tâches et les fichiers de commande.</span><span class="sxs-lookup"><span data-stu-id="642a2-130">It also lists the locations of other files your add-in might make use of, such as the task pane files and command files.</span></span> |

### <a name="script-file"></a><span data-ttu-id="642a2-131">Fichier de script</span><span class="sxs-lookup"><span data-stu-id="642a2-131">Script file</span></span>

<span data-ttu-id="642a2-132">Le fichier de script (**./src/functions/functions.js** ou **./src/functions/functions.ts**) contient le code qui définit des fonctions personnalisées, des commentaires qui définissent la fonction, et associe les noms des fonctions personnalisées à des objets dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="642a2-132">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts**) contains the code that defines custom functions, comments which define the function, and associates the names of the custom functions to objects in the JSON metadata file.</span></span>

<span data-ttu-id="642a2-133">Le code suivant définit la fonction personnalisée `add`.</span><span class="sxs-lookup"><span data-stu-id="642a2-133">The following code defines the custom function `add`.</span></span> <span data-ttu-id="642a2-134">Les commentaires du code sont utilisés pour générer un fichier de métadonnées JSON décrivant la fonction personnalisée pour Excel.</span><span class="sxs-lookup"><span data-stu-id="642a2-134">The code comments are used to generate a JSON metadata file that describes the custom function to Excel.</span></span> <span data-ttu-id="642a2-135">Le commentaire obligatoire `@customfunction` est déclaré en premier, pour indiquer qu’il s’agit d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="642a2-135">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="642a2-136">Vous pouvez également constater que deux paramètres sont déclarés, `first` et `second`, qui sont suivis de leurs propriétés `description`.</span><span class="sxs-lookup"><span data-stu-id="642a2-136">Additionally, you'll notice two parameters are declared, `first` and `second`, which are followed by their `description` properties.</span></span> <span data-ttu-id="642a2-137">Enfin, une description `returns` est fournie.</span><span class="sxs-lookup"><span data-stu-id="642a2-137">Finally, a `returns` description is given.</span></span> <span data-ttu-id="642a2-138">Pour plus d’informations sur les commentaires requis pour votre fonction personnalisée, voir [Créer des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="642a2-138">For more information about what comments are required for your custom function, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="642a2-139">Le code suivant appelle également `CustomFunctions.associate("ADD", add)` pour associer la fonction `add()` avec son ID dans le fichier de métadonnées JSON `ADD`.</span><span class="sxs-lookup"><span data-stu-id="642a2-139">The following code also calls `CustomFunctions.associate("ADD", add)` to associate the function `add()` with its ID in the JSON metadata file `ADD`.</span></span> <span data-ttu-id="642a2-140">Pour plus d’informations sur l’association de fonctions, voir [Meilleures pratiques des fonctions personnalisées](custom-functions-best-practices.md#associating-function-names-with-json-metadata).</span><span class="sxs-lookup"><span data-stu-id="642a2-140">For more information about associating functions, see [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata).</span></span>

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

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="642a2-141">Notez que le fichier **functions.html** qui régit le chargement du runtime de fonctions personnalisées doit créer un lien vers le CDN actuel pour les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="642a2-141">Note that the **functions.html** file, which governs the loading of the custom functions runtime, must link to the current CDN for custom functions.</span></span> <span data-ttu-id="642a2-142">Les projets préparés avec la version actuelle du générateur Yo Office font référence au CDN correct.</span><span class="sxs-lookup"><span data-stu-id="642a2-142">Projects prepared with the current version of the Yo Office generator reference the correct CDN.</span></span> <span data-ttu-id="642a2-143">Si vous mettez à niveau un projet de fonction personnalisée de mars 2019 ou antérieur, vous devez copier le code ci-dessous dans la page \*\* functions.html\*\*.</span><span class="sxs-lookup"><span data-stu-id="642a2-143">If you are retrofitting a previous custom function project from March 2019 or earlier, you need to copy in the code below to the **functions.html** page.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/beta/hosted/custom-functions-runtime.js" type="text/javascript"></script>
```

### <a name="manifest-file"></a><span data-ttu-id="642a2-144">Fichier manifeste</span><span class="sxs-lookup"><span data-stu-id="642a2-144">Manifest file</span></span>

<span data-ttu-id="642a2-145">Le fichier manifeste XML pour un complément qui définit les fonctions personnalisées (**./manifest.xml** du projet créé par le Générateur de Yo Office) spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers HTML, JavaScript et JSON.</span><span class="sxs-lookup"><span data-stu-id="642a2-145">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> 

<span data-ttu-id="642a2-146">Le marquage XML suivant présente un exemple des éléments`<ExtensionPoint>` et `<Resources>` que vous devez inclure dans le manifeste d’un complément pour activer les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="642a2-146">The following basic XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span> <span data-ttu-id="642a2-147">Si vous utilisez le générateur de Yo Office, vos fichiers de fonction personnalisée générés contiennent un fichier manifeste plus complexe que vous pouvez comparer sur [ce dépôt Github](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="642a2-147">If using the Yo Office generator, your generated custom function files will contain a more complex manifest file, which you can compare on [this Github repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).</span></span>

> [!NOTE] 
> <span data-ttu-id="642a2-148">Les URL spécifiées dans le fichier manifeste pour les fonctions personnalisées de fichiers HTML, JavaScript et JSON doivent avoir le même sous-domaine et être accessibles publiquement.</span><span class="sxs-lookup"><span data-stu-id="642a2-148">The URLs specified in the manifest file for the custom functions JavaScript, JSON, and HTML files must be publicly accessible and have the same subdomain.</span></span>

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="642a2-149">Les fonctions dans Excel sont précédées par l’espace de noms spécifié dans votre fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="642a2-149">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="642a2-150">L’espace de noms d’une fonction vient avant le nom de fonction et les deux sont séparés par un point.</span><span class="sxs-lookup"><span data-stu-id="642a2-150">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="642a2-151">Par exemple, pour appeler la fonction `ADD42` dans la cellule de feuille de calcul Excel, vous saisiriez `=CONTOSO.ADD42`, car `CONTOSO` est l’espace de noms et `ADD42` est le nom de la fonction spécifié dans le fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="642a2-151">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="642a2-152">L’espace de noms est destiné à être utilisé comme identificateur de votre entreprise ou du complément.</span><span class="sxs-lookup"><span data-stu-id="642a2-152">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="642a2-153">Un espace de noms ne peut contenir que des points et des caractères alphanumériques.</span><span class="sxs-lookup"><span data-stu-id="642a2-153">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="642a2-154">Co-création</span><span class="sxs-lookup"><span data-stu-id="642a2-154">Coauthoring</span></span>

<span data-ttu-id="642a2-155">Excel Online et Excel sous Windows avec un abonnement Office 365 vous permettent de co-créer des documents et cette fonctionnalité est disponible avec les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="642a2-155">Excel Online and Excel on Windows with an Office 365 subscription allow you to coauthor documents and this feature works with custom functions.</span></span> <span data-ttu-id="642a2-156">Si votre classeur utilise une fonction personnalisée, votre collègue sera invité à charger le complément de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="642a2-156">If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in.</span></span> <span data-ttu-id="642a2-157">Quand vous avez tous les deux chargé le complément, la fonction personnalisée peut partager les résultats via la co-création.</span><span class="sxs-lookup"><span data-stu-id="642a2-157">Once you both have loaded the add-in, the custom function will share results through coauthoring.</span></span>

<span data-ttu-id="642a2-158">Pour plus d’informations sur la co-création, voir [À propos de la co-création dans Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span><span class="sxs-lookup"><span data-stu-id="642a2-158">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="known-issues"></a><span data-ttu-id="642a2-159">Problèmes connus</span><span class="sxs-lookup"><span data-stu-id="642a2-159">Known issues</span></span>

<span data-ttu-id="642a2-160">Consulter les problèmes connus sur notre[repo GitHub Fonctions Excel Personnalisées](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span><span class="sxs-lookup"><span data-stu-id="642a2-160">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="642a2-161">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="642a2-161">Next steps</span></span>

<span data-ttu-id="642a2-162">Vous voulez essayer les fonctions personnalisées ?</span><span class="sxs-lookup"><span data-stu-id="642a2-162">Want to try out custom functions?</span></span> <span data-ttu-id="642a2-163">Consultez la documentation sur le [démarrage rapide de fonction personnalisée](../quickstarts/excel-custom-functions-quickstart.md) ou le [didacticiel sur les fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="642a2-163">Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.</span></span> 

<span data-ttu-id="642a2-164">Un autre moyen simple d’essayer des fonctions personnalisées consiste à utiliser [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), un complément qui vous permet d’expérimenter des fonctions personnalisées directement dans Excel.</span><span class="sxs-lookup"><span data-stu-id="642a2-164">Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel.</span></span> <span data-ttu-id="642a2-165">Vous pouvez essayer de créer votre propre fonction personnalisée ou utiliser les exemples fournis.</span><span class="sxs-lookup"><span data-stu-id="642a2-165">You can try out creating your own custom function or play with the provided samples.</span></span>

<span data-ttu-id="642a2-166">Êtes-vous prêt à en apprendre davantage sur les capacités des fonctions personnalisées ?</span><span class="sxs-lookup"><span data-stu-id="642a2-166">Ready to read more about the capabilities custom functions?</span></span> <span data-ttu-id="642a2-167">Découvrez une vue d’ensemble de l’[architecture des fonctions personnalisées](custom-functions-architecture.md).</span><span class="sxs-lookup"><span data-stu-id="642a2-167">Learn about an overview of [the custom functions architecture](custom-functions-architecture.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="642a2-168">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="642a2-168">See also</span></span> 
* [<span data-ttu-id="642a2-169">Configuration requise de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="642a2-169">Custom functions requirements</span></span>](custom-functions-requirements.md)
* [<span data-ttu-id="642a2-170">Instructions d’attribution de noms</span><span class="sxs-lookup"><span data-stu-id="642a2-170">Naming guidelines</span></span>](custom-functions-naming.md)
* [<span data-ttu-id="642a2-171">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="642a2-171">Best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="642a2-172">Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="642a2-172">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
