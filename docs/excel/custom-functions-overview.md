---
ms.date: 09/26/2019
description: Créer des fonctions personnalisées dans Excel à l’aide de JavaScript.
title: Créer des fonctions personnalisées dans Excel
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 252ff1badd935dda161f474bb7fefa8e782fd1c4
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814464"
---
# <a name="create-custom-functions-in-excel"></a><span data-ttu-id="1bc34-103">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="1bc34-103">Create custom functions in Excel</span></span> 

<span data-ttu-id="1bc34-104">Les fonctions personnalisées permettent aux développeurs d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="1bc34-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="1bc34-105">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="1bc34-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="1bc34-106">Cet article explique comment créer des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="1bc34-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="1bc34-107">L’image animée suivante montre votre classeur appelant une fonction que vous avez créée avec JavaScript ou Typescript.</span><span class="sxs-lookup"><span data-stu-id="1bc34-107">The following animated image shows your workbook calling a function you've created with JavaScript or Typescript.</span></span> <span data-ttu-id="1bc34-108">Dans cet exemple, la fonction personnalisée `=MYFUNCTION.SPHEREVOLUME` calcule le volume d’une sphère.</span><span class="sxs-lookup"><span data-stu-id="1bc34-108">In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

<span data-ttu-id="1bc34-109">Le code suivant définit la fonction personnalisée `=MYFUNCTION.SPHEREVOLUME`.</span><span class="sxs-lookup"><span data-stu-id="1bc34-109">The following code defines the custom function `=MYFUNCTION.SPHEREVOLUME`.</span></span>

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
> <span data-ttu-id="1bc34-110">La section [problèmes connus](#known-issues)plus loin dans cet article indique les limitations en cours de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="1bc34-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="how-a-custom-function-is-defined-in-code"></a><span data-ttu-id="1bc34-111">Comment une fonction personnalisée est définie dans le code</span><span class="sxs-lookup"><span data-stu-id="1bc34-111">How a custom function is defined in code</span></span>

<span data-ttu-id="1bc34-112">Si vous utilisez le [générateur de Yo Office](https://github.com/OfficeDev/generator-office) pour créer un projet de complément de fonctions personnalisées Excel, vous constaterez qu’il crée des fichiers qui contrôlent totalement vos fonctions, votre volet des tâches et votre complément.</span><span class="sxs-lookup"><span data-stu-id="1bc34-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll find that it creates files which control your functions, your task pane, and your add-in overall.</span></span> <span data-ttu-id="1bc34-113">Nous allons vous concentrer sur les fichiers importants pour les fonctions personnalisées :</span><span class="sxs-lookup"><span data-stu-id="1bc34-113">We'll concentrate on the files that are important to custom functions:</span></span>

| <span data-ttu-id="1bc34-114">File</span><span class="sxs-lookup"><span data-stu-id="1bc34-114">File</span></span> | <span data-ttu-id="1bc34-115">Format de fichier</span><span class="sxs-lookup"><span data-stu-id="1bc34-115">File format</span></span> | <span data-ttu-id="1bc34-116">Description</span><span class="sxs-lookup"><span data-stu-id="1bc34-116">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="1bc34-117">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="1bc34-117">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="1bc34-118">ou</span><span class="sxs-lookup"><span data-stu-id="1bc34-118">or</span></span><br/><span data-ttu-id="1bc34-119">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="1bc34-119">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="1bc34-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="1bc34-120">JavaScript</span></span><br/><span data-ttu-id="1bc34-121">ou</span><span class="sxs-lookup"><span data-stu-id="1bc34-121">or</span></span><br/><span data-ttu-id="1bc34-122">TypeScript</span><span class="sxs-lookup"><span data-stu-id="1bc34-122">TypeScript</span></span> | <span data-ttu-id="1bc34-123">Contient le code qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="1bc34-123">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="1bc34-124">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="1bc34-124">**./src/functions/functions.html**</span></span> | <span data-ttu-id="1bc34-125">HTML</span><span class="sxs-lookup"><span data-stu-id="1bc34-125">HTML</span></span> | <span data-ttu-id="1bc34-126">Fournit une référence&lt;script&gt; au fichier JavaScript qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="1bc34-126">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="1bc34-127">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="1bc34-127">**./manifest.xml**</span></span> | <span data-ttu-id="1bc34-128">XML</span><span class="sxs-lookup"><span data-stu-id="1bc34-128">XML</span></span> | <span data-ttu-id="1bc34-129">Spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers JavaScript et HTML qui figurent plus haut dans ce tableau.</span><span class="sxs-lookup"><span data-stu-id="1bc34-129">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript and HTML files that are listed previously in this table.</span></span> <span data-ttu-id="1bc34-130">Répertorie également les emplacements des autres fichiers que votre complément pourrait utiliser, tels que les fichiers du volet des tâches et les fichiers de commande.</span><span class="sxs-lookup"><span data-stu-id="1bc34-130">It also lists the locations of other files your add-in might make use of, such as the task pane files and command files.</span></span> |

### <a name="script-file"></a><span data-ttu-id="1bc34-131">Fichier de script</span><span class="sxs-lookup"><span data-stu-id="1bc34-131">Script file</span></span>

<span data-ttu-id="1bc34-132">Le fichier de script (**./src/functions/functions.js** ou **./src/functions/functions.ts**) contient le code qui définit des fonctions personnalisées et des commentaires qui définissent la fonction.</span><span class="sxs-lookup"><span data-stu-id="1bc34-132">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts**) contains the code that defines custom functions and comments which define the function.</span></span>

<span data-ttu-id="1bc34-133">Le code suivant définit la fonction personnalisée `add`.</span><span class="sxs-lookup"><span data-stu-id="1bc34-133">The following code defines the custom function `add`.</span></span> <span data-ttu-id="1bc34-134">Les commentaires du code sont utilisés pour générer un fichier de métadonnées JSON décrivant la fonction personnalisée pour Excel.</span><span class="sxs-lookup"><span data-stu-id="1bc34-134">The code comments are used to generate a JSON metadata file that describes the custom function to Excel.</span></span> <span data-ttu-id="1bc34-135">Le commentaire obligatoire `@customfunction` est déclaré en premier, pour indiquer qu’il s’agit d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="1bc34-135">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="1bc34-136">Vous pouvez également constater que deux paramètres sont déclarés, `first` et `second`, qui sont suivis de leurs propriétés `description`.</span><span class="sxs-lookup"><span data-stu-id="1bc34-136">Additionally, you'll notice two parameters are declared, `first` and `second`, which are followed by their `description` properties.</span></span> <span data-ttu-id="1bc34-137">Enfin, une description `returns` est fournie.</span><span class="sxs-lookup"><span data-stu-id="1bc34-137">Finally, a `returns` description is given.</span></span> <span data-ttu-id="1bc34-138">Pour plus d’informations sur les commentaires requis pour votre fonction personnalisée, voir [Créer des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="1bc34-138">For more information about what comments are required for your custom function, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

<span data-ttu-id="1bc34-139">Notez que le fichier **functions.html** qui régit le chargement du runtime de fonctions personnalisées doit créer un lien vers le CDN actuel pour les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="1bc34-139">Note that the **functions.html** file, which governs the loading of the custom functions runtime, must link to the current CDN for custom functions.</span></span> <span data-ttu-id="1bc34-140">Les projets préparés avec la version actuelle du générateur Yo Office font référence au CDN correct.</span><span class="sxs-lookup"><span data-stu-id="1bc34-140">Projects prepared with the current version of the Yo Office generator reference the correct CDN.</span></span> <span data-ttu-id="1bc34-141">Si vous mettez à niveau un projet de fonction personnalisée de mars 2019 ou antérieur, vous devez copier le code ci-dessous dans la page \*\* functions.html\*\*.</span><span class="sxs-lookup"><span data-stu-id="1bc34-141">If you are retrofitting a previous custom function project from March 2019 or earlier, you need to copy in the code below to the **functions.html** page.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/custom-functions-runtime.js" type="text/javascript"></script>
```

### <a name="manifest-file"></a><span data-ttu-id="1bc34-142">Fichier manifeste</span><span class="sxs-lookup"><span data-stu-id="1bc34-142">Manifest file</span></span>

<span data-ttu-id="1bc34-143">Le fichier manifeste XML pour un complément qui définit les fonctions personnalisées (**./manifest.xml** du projet créé par le Générateur de Yo Office) spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers HTML, JavaScript et JSON.</span><span class="sxs-lookup"><span data-stu-id="1bc34-143">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span>

<span data-ttu-id="1bc34-144">Le marquage XML suivant présente un exemple des éléments`<ExtensionPoint>` et `<Resources>` que vous devez inclure dans le manifeste d’un complément pour activer les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="1bc34-144">The following basic XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span> <span data-ttu-id="1bc34-145">Si vous utilisez le générateur de Yo Office, vos fichiers de fonction personnalisée générés contiennent un fichier manifeste plus complexe que vous pouvez comparer sur [ce dépôt Github](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="1bc34-145">If using the Yo Office generator, your generated custom function files will contain a more complex manifest file, which you can compare on [this Github repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml).</span></span>

> [!NOTE] 
> <span data-ttu-id="1bc34-146">Les URL spécifiées dans le fichier manifeste pour les fonctions personnalisées de fichiers HTML, JavaScript et JSON doivent avoir le même sous-domaine et être accessibles publiquement.</span><span class="sxs-lookup"><span data-stu-id="1bc34-146">The URLs specified in the manifest file for the custom functions JavaScript, JSON, and HTML files must be publicly accessible and have the same subdomain.</span></span>

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
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
> <span data-ttu-id="1bc34-147">Les fonctions dans Excel sont précédées par l’espace de noms spécifié dans votre fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="1bc34-147">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="1bc34-148">L’espace de noms d’une fonction vient avant le nom de fonction et les deux sont séparés par un point.</span><span class="sxs-lookup"><span data-stu-id="1bc34-148">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="1bc34-149">Par exemple, pour appeler la fonction `ADD42` dans la cellule de feuille de calcul Excel, vous saisiriez `=CONTOSO.ADD42`, car `CONTOSO` est l’espace de noms et `ADD42` est le nom de la fonction spécifié dans le fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="1bc34-149">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="1bc34-150">L’espace de noms est destiné à être utilisé comme identificateur de votre entreprise ou du complément.</span><span class="sxs-lookup"><span data-stu-id="1bc34-150">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="1bc34-151">Un espace de noms ne peut contenir que des points et des caractères alphanumériques.</span><span class="sxs-lookup"><span data-stu-id="1bc34-151">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="1bc34-152">Co-création</span><span class="sxs-lookup"><span data-stu-id="1bc34-152">Coauthoring</span></span>

<span data-ttu-id="1bc34-153">Excel sur le web et Windows avec un abonnement Office 365 vous permettent de co-créer des documents et cette fonctionnalité est disponible avec les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="1bc34-153">Excel on the web and Windows connected to an Office 365 subscription allow you to coauthor documents and this feature works with custom functions.</span></span> <span data-ttu-id="1bc34-154">Si votre classeur utilise une fonction personnalisée, votre collègue sera invité à charger le complément de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="1bc34-154">If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in.</span></span> <span data-ttu-id="1bc34-155">Quand vous avez tous les deux chargé le complément, la fonction personnalisée peut partager les résultats via la co-création.</span><span class="sxs-lookup"><span data-stu-id="1bc34-155">Once you both have loaded the add-in, the custom function will share results through coauthoring.</span></span>

<span data-ttu-id="1bc34-156">Pour plus d’informations sur la co-création, voir [À propos de la co-création dans Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span><span class="sxs-lookup"><span data-stu-id="1bc34-156">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="known-issues"></a><span data-ttu-id="1bc34-157">Problèmes connus</span><span class="sxs-lookup"><span data-stu-id="1bc34-157">Known issues</span></span>

<span data-ttu-id="1bc34-158">Consulter les problèmes connus sur notre[repo GitHub Fonctions Excel Personnalisées](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span><span class="sxs-lookup"><span data-stu-id="1bc34-158">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="1bc34-159">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="1bc34-159">Next steps</span></span>

<span data-ttu-id="1bc34-160">Vous voulez essayer les fonctions personnalisées ?</span><span class="sxs-lookup"><span data-stu-id="1bc34-160">Want to try out custom functions?</span></span> <span data-ttu-id="1bc34-161">Consultez la documentation sur le [démarrage rapide de fonction personnalisée](../quickstarts/excel-custom-functions-quickstart.md) ou le [didacticiel sur les fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="1bc34-161">Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.</span></span>

<span data-ttu-id="1bc34-162">Un autre moyen simple d’essayer des fonctions personnalisées consiste à utiliser [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), un complément qui vous permet d’expérimenter des fonctions personnalisées directement dans Excel.</span><span class="sxs-lookup"><span data-stu-id="1bc34-162">Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel.</span></span> <span data-ttu-id="1bc34-163">Vous pouvez essayer de créer votre propre fonction personnalisée ou utiliser les exemples fournis.</span><span class="sxs-lookup"><span data-stu-id="1bc34-163">You can try out creating your own custom function or play with the provided samples.</span></span>

<span data-ttu-id="1bc34-164">Êtes-vous prêt à en apprendre davantage sur les capacités des fonctions personnalisées ?</span><span class="sxs-lookup"><span data-stu-id="1bc34-164">Ready to read more about the capabilities custom functions?</span></span> <span data-ttu-id="1bc34-165">Découvrez une vue d’ensemble de l’[architecture des fonctions personnalisées](custom-functions-architecture.md).</span><span class="sxs-lookup"><span data-stu-id="1bc34-165">Learn about an overview of [the custom functions architecture](custom-functions-architecture.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="1bc34-166">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="1bc34-166">See also</span></span> 
* [<span data-ttu-id="1bc34-167">Configuration requise de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="1bc34-167">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="1bc34-168">Instructions d’attribution de noms</span><span class="sxs-lookup"><span data-stu-id="1bc34-168">Naming guidelines</span></span>](custom-functions-naming.md)
* [<span data-ttu-id="1bc34-169">Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="1bc34-169">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
