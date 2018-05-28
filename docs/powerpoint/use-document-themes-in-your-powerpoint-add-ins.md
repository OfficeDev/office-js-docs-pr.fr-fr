---
title: Utiliser des th?mes de document dans vos compl?ments PowerPoint
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: e39f323f842112970d8e2a7473fa5db9dca0e55c
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="use-document-themes-in-your-powerpoint-add-ins"></a><span data-ttu-id="88436-102">Utiliser des th?mes de document dans vos compl?ments PowerPoint</span><span class="sxs-lookup"><span data-stu-id="88436-102">Use document themes in your PowerPoint add-ins</span></span>

<span data-ttu-id="88436-p101">Un [th?me Office](https://support.office.com/en-US/Article/What-is-a-theme--7528ccc2-4327-4692-8bf5-9b5a3f2a5ef5) est constitu?, en partie, d?un jeu de polices et de couleurs visuellement assortis que vous pouvez appliquer ? des pr?sentations, des documents, des feuilles de calcul et des courriers ?lectroniques. Pour appliquer ou personnaliser le th?me d?une pr?sentation dans PowerPoint, utilisez les groupes **Th?mes** et **Variantes** dans l?onglet **Conception** du ruban. PowerPoint affecte le **th?me Office** par d?faut ? chaque nouvelle pr?sentation vierge, mais vous pouvez choisir parmi les autres th?mes disponibles dans l?onglet **Conception**, t?l?charger des th?mes suppl?mentaires ? partir d?Office.com, ou cr?er et personnaliser votre propre th?me.</span><span class="sxs-lookup"><span data-stu-id="88436-p101">An [Office theme](https://support.office.com/en-US/Article/What-is-a-theme--7528ccc2-4327-4692-8bf5-9b5a3f2a5ef5) consists, in part, of a visually coordinated set of fonts and colors that you can apply to presentations, documents, worksheets, and emails. To apply or customize the theme of a presentation in PowerPoint, you use the **Themes** and **Variants** groups on **Design** tab of the ribbon. PowerPoint assigns a new blank presentation with the default **Office Theme**, but you can choose other themes available on the **Design** tab, download additional themes from Office.com, or create and customize your own theme.</span></span>

<span data-ttu-id="88436-106">OfficeThemes.css vous permet de concevoir des compl?ments coordonn?s ? PowerPoint de deux fa?ons :</span><span class="sxs-lookup"><span data-stu-id="88436-106">Using OfficeThemes.css, helps you design add-ins that are coordinated with PowerPoint in two ways:</span></span>

- <span data-ttu-id="88436-p102">**Dans les compl?ments de contenu pour PowerPoint**. Utilisez les classes de th?mes du document d?OfficeThemes.css pour sp?cifier les polices et les couleurs correspondant au th?me de la pr?sentation dans laquelle votre contenu compl?ment est ins?r? ; ces polices et couleurs seront mises ? jour dynamiquement si un utilisateur modifie ou personnalise le th?me de la pr?sentation.</span><span class="sxs-lookup"><span data-stu-id="88436-p102">**In content add-ins for PowerPoint**. Use the document theme classes of OfficeThemes.css to specify fonts and colors that match the theme of the presentation your content add-in is inserted into - and those fonts and colors will dynamically update if a user changes or customizes the presentation's theme.</span></span>
    
- <span data-ttu-id="88436-p103">**Dans les compl?ments du volet Office pour PowerPoint**. Utilisez les classes de th?mes de l?interface utilisateur Office d?OfficeThemes.css pour sp?cifier les m?mes polices et couleurs d?arri?re-plan que celles utilis?es dans l?interface utilisateur, de sorte que vos compl?ments du volet Office correspondent aux couleurs des volets Office int?gr?s ; ces couleurs seront mises ? jour dynamiquement si un utilisateur modifie le th?me de l?interface utilisateur Office.</span><span class="sxs-lookup"><span data-stu-id="88436-p103">**In task pane add-ins for PowerPoint**. Use the Office UI theme classes of OfficeThemes.css to specify the same fonts and background colors used in the UI so that your task pane add-ins will match the colors of built-in task panes - and those colors will dynamically update if a user changes the Office UI theme.</span></span>

### <a name="document-theme-colors"></a><span data-ttu-id="88436-111">Couleurs de th?me de document</span><span class="sxs-lookup"><span data-stu-id="88436-111">Document theme colors</span></span>

<span data-ttu-id="88436-p104">Chaque th?me de document Office d?finit 12 couleurs. Dix de ces couleurs sont disponibles lorsque vous d?finissez la police, l?arri?re-plan et d?autres param?tres de couleur dans une pr?sentation gr?ce au s?lecteur de couleurs.</span><span class="sxs-lookup"><span data-stu-id="88436-p104">Every Office document theme defines 12 colors. Ten of these colors are available when you set font, background, and other color settings in a presentation with the color picker.</span></span>

![Palette de couleurs](../images/office15-app-color-palette.png)

<span data-ttu-id="88436-115">Pour afficher ou personnaliser l?int?gralit? des 12 couleurs de th?me dans PowerPoint, dans le groupe **Variantes** de l?onglet **Conception**, cliquez sur le menu d?roulant **Plus**, puis choisissez **Couleur** et cliquez sur **Personnaliser les couleurs** pour afficher la bo?te de dialogue **Cr?er de nouvelles couleurs de th?me**.</span><span class="sxs-lookup"><span data-stu-id="88436-115">To view or customize the full set of 12 theme colors in PowerPoint, in the  **Variants** group on the **Design** tab, click the **More** drop-down - then point to **Color**, and click  **Customize Colors** to display the **Create New Theme Colors** dialog box.</span></span>

![Bo?te de dialogue Cr?er de nouvelles couleurs de th?me](../images/office15-app-create-new-theme-colors.png)

<span data-ttu-id="88436-p105">Les quatre premi?res couleurs sont pour le texte et les arri?re-plans. Un texte cr?? avec des couleurs claires sera toujours lisible sur les couleurs fonc?es, tandis qu?un texte cr?? avec des couleurs fonc?es sera toujours lisible sur les couleurs claires. Les six couleurs suivantes sont des couleurs d?accentuation qui sont toujours visibles sur les quatre couleurs d?arri?re-plan potentielles. Les deux derni?res couleurs sont pour les liens hypertexte et les liens hypertexte visit?s.</span><span class="sxs-lookup"><span data-stu-id="88436-p105">The first four colors are for text and backgrounds. Text that is created with the light colors will always be legible over the dark colors, and text that is created with dark colors will always be legible over the light colors. The next six are accent colors that are always visible over the four potential background colors. The last two colors are for hyperlinks and followed hyperlinks.</span></span>

### <a name="document-theme-fonts"></a><span data-ttu-id="88436-121">Polices de th?me de document</span><span class="sxs-lookup"><span data-stu-id="88436-121">Document theme fonts</span></span>

<span data-ttu-id="88436-p106">Chaque th?me de document Office d?finit ?galement deux polices : une pour les titres et l?autre pour le corps de texte. PowerPoint utilise ces polices pour cr?er des styles de texte automatiques. En outre, les galeries **Styles rapides** pour le texte et **WordArt** utilisent ces m?mes polices de th?me. Ces deux polices sont les deux premi?res propos?es lorsque vous s?lectionnez des polices avec le s?lecteur de polices.</span><span class="sxs-lookup"><span data-stu-id="88436-p106">Every Office document theme also defines two fonts -- one for headings and one for body text. PowerPoint uses these fonts to construct automatic text styles. In addition,  **Quick Styles** galleries for text and **WordArt** use these same theme fonts. These two fonts are available as the first two selections when you select fonts with the font picker.</span></span>

![S?lecteur de polices](../images/office15-app-font-picker.png)

<span data-ttu-id="88436-127">Pour afficher ou personnaliser les polices de th?me dans PowerPoint, dans le groupe **Variantes** de l?onglet **Conception**, cliquez sur le menu d?roulant **Plus**. Ensuite, pointez vers **Polices** et cliquez sur **Personnaliser les polices** pour afficher la bo?te de dialogue **Cr?er de nouvelles polices de th?me**.</span><span class="sxs-lookup"><span data-stu-id="88436-127">To view or customize theme fonts in PowerPoint, in the  **Variants** group on the **Design** tab, click the **More** drop-down - then point to **Fonts**, and click  **Customize Fonts** to display the **Create New Theme Fonts** dialog box.</span></span>

![Bo?te de dialogue Cr?er de nouvelles polices de th?me](../images/office15-app-create-new-theme-fonts.png)

### <a name="office-ui-theme-fonts-and-colors"></a><span data-ttu-id="88436-129">Couleurs et polices de th?me de l?interface utilisateur Office</span><span class="sxs-lookup"><span data-stu-id="88436-129">Office UI theme fonts and colors</span></span>

<span data-ttu-id="88436-p107">Office vous permet ?galement de choisir entre plusieurs th?mes pr?d?finis qui sp?cifient certaines des couleurs et des polices utilis?es dans l?interface utilisateur de toutes les applications Office. Pour cela, utilisez le menu d?roulant  **Fichier**  >   **Compte**  >   **Th?me Office** (dans toutes les applications Office).</span><span class="sxs-lookup"><span data-stu-id="88436-p107">Office also lets you choose between several predefined themes that specify some of the colors and fonts used in the UI of all Office applications. To do that, you use the  **File** > **Account** > **Office Theme** drop-down (from any Office application).</span></span>

![Liste d?roulante de th?mes Office](../images/office15-app-office-theme-picker.png)

<span data-ttu-id="88436-p108">OfficeThemes.css inclut des classes que vous pouvez utiliser dans vos compl?ments du volet Office pour PowerPoint afin qu?elles utilisent ces m?mes polices et couleurs. Cela vous permet de concevoir des compl?ments du volet Office dont l?apparence concorde avec celle des volets Office int?gr?s.</span><span class="sxs-lookup"><span data-stu-id="88436-p108">OfficeThemes.css includes classes that you can use in your task pane add-ins for PowerPoint so they will use these same fonts and colors. This lets you design your task pane add-ins that match the appearance of built-in task panes.</span></span>

## <a name="using-officethemescss"></a><span data-ttu-id="88436-135">Utilisation d?OfficeThemes.css</span><span class="sxs-lookup"><span data-stu-id="88436-135">Using OfficeThemes.css</span></span>

<span data-ttu-id="88436-p109">En utilisant le fichier OfficeThemes.css avec vos compl?ments de contenu pour PowerPoint, vous pouvez coordonner l?apparence de votre compl?ment avec le th?me appliqu? ? la pr?sentation avec laquelle elle est ex?cut?e. En utilisant le fichier OfficeThemes.css avec vos compl?ments du volet Office pour PowerPoint, vous pouvez coordonner l?apparence de votre compl?ment avec les polices et couleurs de l?interface utilisateur Office.</span><span class="sxs-lookup"><span data-stu-id="88436-p109">Using the OfficeThemes.css file with your content add-ins for PowerPoint lets you coordinate the appearance of your add-in with the theme applied to the presentation it's running with. Using the OfficeThemes.css file with your task pane add-ins for PowerPoint lets you coordinate the appearance of your add-in with the fonts and colors of the Office UI.</span></span>

### <a name="adding-the-officethemescss-file-to-your-project"></a><span data-ttu-id="88436-138">Ajout du fichier OfficeThemes.css ? votre projet</span><span class="sxs-lookup"><span data-stu-id="88436-138">Adding the OfficeThemes.css file to your project</span></span>

<span data-ttu-id="88436-139">Suivez la proc?dure suivante pour ajouter et r?f?rencer le fichier OfficeThemes.css dans votre projet compl?ment.</span><span class="sxs-lookup"><span data-stu-id="88436-139">Use the following steps to add and reference the OfficeThemes.css file to your add-in project.</span></span>

#### <a name="to-add-officethemescss-to-your-visual-studio-project"></a><span data-ttu-id="88436-140">Pour ajouter le fichier OfficeThemes.css ? votre projet Visual Studio</span><span class="sxs-lookup"><span data-stu-id="88436-140">To add OfficeThemes.css to your Visual Studio project</span></span>

1. <span data-ttu-id="88436-141">Dans l?**explorateur de solutions**, cliquez avec le bouton droit de la souris sur le dossier **Contenu** dans le projet _**project_name**_**Web**, pointez sur **Ajouter** et s?lectionnez **Feuille de style**.</span><span class="sxs-lookup"><span data-stu-id="88436-141">In **Solution Explorer**, right-click the **Content** folder in the _**project_name**_**Web** project, point to **Add**, and then select **Style Sheet**.</span></span>
    
2. <span data-ttu-id="88436-142">Nommez la nouvelle feuille de style **OfficeThemes**.</span><span class="sxs-lookup"><span data-stu-id="88436-142">Name the new style sheet **OfficeThemes**.</span></span>
    
   > [!IMPORTANT]
   > <span data-ttu-id="88436-143">Le nom de la feuille de style doit ?tre OfficeThemes, sinon la fonctionnalit? qui met ? jour dynamiquement les polices et couleurs de compl?ment lorsqu?un utilisateur modifie le th?me ne fonctionnera pas.</span><span class="sxs-lookup"><span data-stu-id="88436-143">The style sheet must be named OfficeThemes, or the feature that dynamically updates add-in fonts and colors when a user changes the theme won't work.</span></span>
   
3. <span data-ttu-id="88436-144">Supprimez la classe **body** par d?faut (`body {}`) dans le fichier, et copiez-collez le code CSS suivant dans le fichier.</span><span class="sxs-lookup"><span data-stu-id="88436-144">Delete the default **body** class (`body {}`) in the file, and copy and paste the following CSS code into the file.</span></span>
    
    ```css
    /* The following classes describe the common theme information for office documents */ 

    /* Basic Font and Background Colors for text */ 
    .office-docTheme-primary-fontColor { color:#000000; } 
    .office-docTheme-primary-bgColor { background-color:#ffffff; } 
    .office-docTheme-secondary-fontColor { color: #000000; } 
    .office-docTheme-secondary-bgColor { background-color: #ffffff; } 

    /* Accent color definitions for fonts */ 
    .office-contentAccent1-color { color:#5b9bd5; } 
    .office-contentAccent2-color { color:#ed7d31; } 
    .office-contentAccent3-color { color:#a5a5a5; } 
    .office-contentAccent4-color { color:#ffc000; } 
    .office-contentAccent5-color { color:#4472c4; } 
    .office-contentAccent6-color { color:#70ad47; } 

    /* Accent color for backgrounds */ 
    .office-contentAccent1-bgColor { background-color:#5b9bd5; } 
    .office-contentAccent2-bgColor { background-color:#ed7d31; } 
    .office-contentAccent3-bgColor { background-color:#a5a5a5; } 
    .office-contentAccent4-bgColor { background-color:#ffc000; } 
    .office-contentAccent5-bgColor { background-color:#4472c4; } 
    .office-contentAccent6-bgColor { background-color:#70ad47; } 

    /* Accent color for borders */ 
    .office-contentAccent1-borderColor { border-color:#5b9bd5; } 
    .office-contentAccent2-borderColor { border-color:#ed7d31; } 
    .office-contentAccent3-borderColor { border-color:#a5a5a5; } 
    .office-contentAccent4-borderColor { border-color:#ffc000; } 
    .office-contentAccent5-borderColor { border-color:#4472c4; } 
    .office-contentAccent6-borderColor { border-color:#70ad47; } 

    /* links */ 
    .office-a { color: #0563c1; } 
    .office-a:visited { color: #954f72; } 

    /* Body Fonts */ 
    .office-bodyFont-eastAsian { } /* East Asian name of the Font */ 
    .office-bodyFont-latin { font-family:"Calibri"; } /* Latin name of the Font */ 
    .office-bodyFont-script { } /* Script name of the Font */ 
    .office-bodyFont-localized { font-family:"Calibri"; } /* Localized name of the Font. Corresponds to the default font of the culture currently used in Office.*/ 

    /* Headers Font */ 
    .office-headerFont-eastAsian { } 
    .office-headerFont-latin { font-family:"Calibri Light"; } 
    .office-headerFont-script { } 
    .office-headerFont-localized { font-family:"Calibri Light"; } 

    /* The following classes define font and background colors for Office UI themes. These classes should only be used in task pane add-ins */ 

    /* Basic Font and Background Colors for PPT */ 
    .office-officeTheme-primary-fontColor { color:#b83b1d; } 
    .office-officeTheme-primary-bgColor { background-color:#dedede; } 
    .office-officeTheme-secondary-fontColor { color:#262626; } 
    .office-officeTheme-secondary-bgColor { background-color:#ffffff; }
    ```
4. <span data-ttu-id="88436-145">Si vous utilisez un autre outil que Visual Studio pour cr?er votre compl?ment, copiez le code CSS de l??tape 3 dans un fichier texte, en vous assurant que le fichier est enregistr? sous le nom OfficeThemes.css.</span><span class="sxs-lookup"><span data-stu-id="88436-145">If you are using a tool other than Visual Studio to create your add-in, copy the CSS code from step 3 into a text file, making sure to save the file as OfficeThemes.css.</span></span>   

### <a name="referencing-officethemescss-in-your-add-ins-html-pages"></a><span data-ttu-id="88436-146">R?f?rencement d?OfficeThemes.css dans les pages HTML de votre compl?ment</span><span class="sxs-lookup"><span data-stu-id="88436-146">Referencing OfficeThemes.css in your add-in's HTML pages</span></span>

<span data-ttu-id="88436-147">Pour utiliser le fichier OfficeThemes.css dans votre projet de compl?ment, ajoutez une balise `<link>` r?f?ren?ant le fichier OfficeThemes.css ? l?int?rieur de la balise `<head>` des pages web (par exemple, un fichier .html, .aspx ou .php) qui impl?mentent l?interface utilisateur de votre compl?ment au format suivant :</span><span class="sxs-lookup"><span data-stu-id="88436-147">To use the OfficeThemes.css file in your add-in project, add a `<link>` tag that references the OfficeThemes.css file inside the `<head>` tag of the web pages (such as an .html, .aspx, or .php file) that implement the UI of your add-in in this format:</span></span>

```HTML
<link href="<local_path_to_OfficeThemes.css>" rel="stylesheet" type="text/css" />
```

<span data-ttu-id="88436-148">Pour effectuer cette op?ration dans Visual Studio, proc?dez comme suit.</span><span class="sxs-lookup"><span data-stu-id="88436-148">To do this in Visual Studio, follow these steps.</span></span>

#### <a name="to-reference-officethemescss-in-your-add-in-for-powerpoint"></a><span data-ttu-id="88436-149">Pour r?f?rencer OfficeThemes.css dans votre compl?ment PowerPoint</span><span class="sxs-lookup"><span data-stu-id="88436-149">To reference OfficeThemes.css in your add-in for PowerPoint</span></span>

1. <span data-ttu-id="88436-150">Dans Visual Studio 2015, ouvrez ou cr?ez un projet de **compl?ment Office**.</span><span class="sxs-lookup"><span data-stu-id="88436-150">In Visual Studio 2015, open or create a new **Office Add-in** project.</span></span>
    
2. <span data-ttu-id="88436-151">Dans les pages HTML qui impl?mentent l?interface utilisateur de votre compl?ment, telles que Home.html dans le mod?le par d?faut, ajoutez la balise `<link>` suivante ? l?int?rieur de la balise `<head>` qui r?f?rence le fichier OfficeThemes.css :</span><span class="sxs-lookup"><span data-stu-id="88436-151">In the HTML pages that implement the UI of your add-in, such as Home.html in the default template, add the following `<link>` tag inside the `<head>` tag that references the OfficeThemes.css file:</span></span>
    
    ```HTML
    <link href="../../Content/OfficeThemes.css" rel="stylesheet" type="text/css" />
    ```

<span data-ttu-id="88436-152">Si vous cr?ez votre compl?ment avec un outil autre que Visual Studio, ajoutez une balise `<link>` avec le m?me format sp?cifiant un chemin d?acc?s relatif vers la copie d?OfficeThemes.css qui sera d?ploy?e avec votre compl?ment.</span><span class="sxs-lookup"><span data-stu-id="88436-152">If you are creating your add-in with a tool other than Visual Studio, add a `<link>` tag with the same format specifying a relative path to the copy of OfficeThemes.css that will be deployed with your add-in.</span></span>

### <a name="using-officethemescss-document-theme-classes-in-your-content-add-ins-html-page"></a><span data-ttu-id="88436-153">Utilisation de classes de th?mes de document OfficeThemes.css dans la page HTML de votre compl?ment de contenu</span><span class="sxs-lookup"><span data-stu-id="88436-153">Using OfficeThemes.css document theme classes in your content add-in's HTML page</span></span>

<span data-ttu-id="88436-p110">Ci-dessous figure un exemple simple de code HTML dans une compl?ment de contenu qui utilise les classes de th?mes de document OfficeTheme.css. Pour plus d?informations sur les classes OfficeThemes.css qui correspondent aux 12 couleurs et aux 2 polices utilis?es dans un th?me de document, voir [Classes de th?mes pour les compl?ments de contenu](#theme-classes-for-content-add-ins).</span><span class="sxs-lookup"><span data-stu-id="88436-p110">The following shows a simple example of HTML in a content add-in that uses the OfficeTheme.css document theme classes. For details about the OfficeThemes.css classes that correspond to the 12 colors and 2 fonts used in a document theme, see [Theme classes for content add-ins](#theme-classes-for-content-add-ins).</span></span>

```HTML
<body>
    <div id="themeSample" class="office-docTheme-primary-fontColor ">
        <h1 class="office-headerFont-latin">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent1-bgColor">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent2-bgColor">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent3-bgColor">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent4-bgColor">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent5-bgColor">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent6-bgColor">Hello world!</h1> 
        <p class="office-bodyFont-latin office-docTheme-secondary-fontColor">Hello world!</p> 
    </div>
</body>
```

<span data-ttu-id="88436-156">Lors de l?ex?cution, lorsque le compl?ment de contenu est ins?r? dans une pr?sentation qui utilise le **th?me Office** par d?faut, il est restitu? comme suit.</span><span class="sxs-lookup"><span data-stu-id="88436-156">At runtime, when inserted into a presentation that uses the default  **Office Theme**, the content add-in is rendered like this.</span></span>

![Application de contenu en cours d?ex?cution avec le th?me Office](../images/office15-app-content-app-office-theme.png)

<span data-ttu-id="88436-p111">Si vous modifiez la pr?sentation afin d?utiliser un autre th?me ou de personnaliser le th?me de la pr?sentation, les polices et couleurs sp?cifi?es avec des classes OfficeThemes.css sont mises ? jour dynamiquement pour correspondre aux polices et aux couleurs du th?me de la pr?sentation. En prenant l?exemple HTML ci-dessus, si la pr?sentation dans laquelle le compl?ment est ins?r? utilise le th?me **Facette**, le compl?ment est restitu? comme suit.</span><span class="sxs-lookup"><span data-stu-id="88436-p111">If you change the presentation to use another theme or customize the presentation's theme, the fonts and colors specified with OfficeThemes.css classes will dynamically update to correspond to the fonts and colors of the presentation's theme. Using the same HTML example as above, if the presentation the add-in is inserted into uses the **Facet** theme, the add-in rendering will look like this.</span></span>

![Application de contenu en cours d?ex?cution avec le th?me Facette](../images/office15-app-content-app-facet-theme.png)


### <a name="using-officethemescss-office-ui-theme-classes-in-your-task-pane-add-ins-html-page"></a><span data-ttu-id="88436-161">Utilisation de classes de th?mes de l?interface utilisateur Office OfficeThemes.css dans la page HTML de votre compl?ment du volet Office</span><span class="sxs-lookup"><span data-stu-id="88436-161">Using OfficeThemes.css Office UI theme classes in your task pane add-in's HTML page</span></span>

<span data-ttu-id="88436-162">Outre le th?me du document, les utilisateurs peuvent personnaliser le mod?le de couleurs de l?interface utilisateur Office de toutes les applications Office ? l?aide de la zone de liste d?roulante **Fichier** > **Compte** > **Th?me Office**.</span><span class="sxs-lookup"><span data-stu-id="88436-162">In addition to the document theme, users can customize the color scheme of the Office user interface for all Office applications using the **File** > **Account** > **Office Theme** drop-down box.</span></span>

<span data-ttu-id="88436-p112">Ci-dessous figure un exemple simple de code HTML dans une compl?ment de volet Office qui utilise des classes OfficeTheme.css pour sp?cifier les couleurs de police et d?arri?re-plan. Pour plus d?informations sur les classes OfficeThemes.css qui correspondent aux polices et aux couleurs du th?me de l?interface utilisateur Office, voir [Classes de th?mes pour les compl?ments du volet Office](#theme-classes-for-task-pane-add-ins).</span><span class="sxs-lookup"><span data-stu-id="88436-p112">The following shows a simple example of HTML in a task pane add-in that uses OfficeTheme.css classes to specify font color and background color. For details about the OfficeThemes.css classes that correspond to fonts and colors of the Office UI theme, see [Theme classes for task pane add-ins](#theme-classes-for-task-pane-add-ins).</span></span>

```HTML
<body> 
    <div id="content-header" class="office-officeTheme-primary-fontColor office-officeTheme-primary-bgColor"> 
        <div class="padding">
            <h1>Welcome</h1>
        </div> 
    </div> 
    <div id="content-main" class="office-officeTheme-secondary-fontColor office-officeTheme-secondary-bgColor"> 
        <div class="padding"> 
            <p>Add home screen content here.</p> 
            <p>For example:</p> 
            <button id="get-data-from-selection">Get data from selection</button> 
            <p><a target="_blank" class="office-a" href="https://go.microsoft.com/fwlink/?LinkId=276812">Find more samples online...</a></p>
        </div>
    </div>
</body> 
```

<br/>

<span data-ttu-id="88436-165">Lors de l?ex?cution de PowerPoint avec **Fichier** > **Compte** > **Th?me Office** d?fini sur **Blanc**, le compl?ment de volet de t?ches est restitu? comme suit.</span><span class="sxs-lookup"><span data-stu-id="88436-165">When running in PowerPoint with **File** > **Account** > **Office Theme** set to **White**, the task pane add-in is rendered like this.</span></span>

![Volet de t?ches avec th?me blanc Office](../images/office15-app-task-pane-theme-white.png)

<br/>

<span data-ttu-id="88436-167">Si vous modifiez la valeur de **Th?me Office** en la d?finissant sur **Gris fonc?**, les polices et couleurs sp?cifi?es avec des classes OfficeThemes.css seront mises ? jour dynamiquement et seront restitu?es comme suit.</span><span class="sxs-lookup"><span data-stu-id="88436-167">If you change **OfficeTheme** to **Dark Gray**, the fonts and colors specified with OfficeThemes.css classes will dynamically update to render like this.</span></span>

![Volet de t?ches avec th?me gris fonc? Office](../images/office15-app-task-pane-theme-dark-gray.png)

<br/>

## <a name="officethemecss-classes"></a><span data-ttu-id="88436-169">Classes OfficeTheme.css</span><span class="sxs-lookup"><span data-stu-id="88436-169">OfficeTheme.css classes</span></span>

<span data-ttu-id="88436-170">Le fichier OfficeThemes.css contient deux jeux de classes que vous pouvez utiliser avec vos compl?ments de contenu et du volet Office PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="88436-170">The OfficeThemes.css file contains two sets of classes you can use with your content and task pane add-ins for PowerPoint.</span></span>

### <a name="theme-classes-for-content-add-ins"></a><span data-ttu-id="88436-171">Classes de th?mes pour les compl?ments de contenu</span><span class="sxs-lookup"><span data-stu-id="88436-171">Theme classes for content add-ins</span></span>

<span data-ttu-id="88436-p113">Le fichier OfficeThemes.css fournit des classes qui correspondent aux 12 couleurs et aux 2 polices utilis?es dans un th?me de document. Ces classes sont adapt?es aux compl?ments de contenu pour PowerPoint, de sorte que les polices et les couleurs de votre compl?ment seront en harmonie avec la pr?sentation dans laquelle votre compl?ment est ins?r?.</span><span class="sxs-lookup"><span data-stu-id="88436-p113">The OfficeThemes.css file provides classes that correspond to the 2 fonts and 12 colors used in a document theme. These classes are appropriate to use with content add-ins for PowerPoint so that your add-in's fonts and colors will be coordinated with the presentation it's inserted into.</span></span>

#### <a name="theme-fonts-for-content-add-ins"></a><span data-ttu-id="88436-174">Polices de th?me pour les compl?ments de contenu</span><span class="sxs-lookup"><span data-stu-id="88436-174">Theme fonts for content add-ins</span></span>

|<span data-ttu-id="88436-175">**Classe**</span><span class="sxs-lookup"><span data-stu-id="88436-175">**Class**</span></span>|<span data-ttu-id="88436-176">**Description**</span><span class="sxs-lookup"><span data-stu-id="88436-176">**Description**</span></span>|
|:-----|:-----|
| `office-bodyFont-eastAsian`|<span data-ttu-id="88436-177">Nom en langues d?Asie de l?Est de la police du corps de texte.</span><span class="sxs-lookup"><span data-stu-id="88436-177">East Asian name of the body font.</span></span>|
| `office-bodyFont-latin`|<span data-ttu-id="88436-p114">Nom latin de la police du corps de texte (par d?faut, ? Calibri ?).</span><span class="sxs-lookup"><span data-stu-id="88436-p114">Latin name of the body font. Default "Calabri"</span></span>|
| `office-bodyFont-script`|<span data-ttu-id="88436-180">Nom de script de la police du corps de texte.</span><span class="sxs-lookup"><span data-stu-id="88436-180">Script name of the body font.</span></span>|
| `office-bodyFont-localized`|<span data-ttu-id="88436-p115">Nom localis? de la police du corps de texte. Sp?cifie le nom de la police par d?faut en fonction de la culture actuellement utilis?e dans Office.</span><span class="sxs-lookup"><span data-stu-id="88436-p115">Localized name of the body font. Specifies the default font name according to the culture currently used in Office.</span></span>|
| `office-headerFont-eastAsian`|<span data-ttu-id="88436-183">Nom en langues d?Asie de l?Est de la police des en-t?tes.</span><span class="sxs-lookup"><span data-stu-id="88436-183">East Asian name of the headers font.</span></span>|
| `office-headerFont-latin`|<span data-ttu-id="88436-p116">Nom latin de la police des en-t?tes (par d?faut, ? Calibri Light ?).</span><span class="sxs-lookup"><span data-stu-id="88436-p116">Latin name of the headers font. Default "Calabri Light"</span></span>|
| `office-headerFont-script`|<span data-ttu-id="88436-186">Nom de script de la police des en-t?tes.</span><span class="sxs-lookup"><span data-stu-id="88436-186">Script name of the headers font.</span></span>|
| `office-headerFont-localized`|<span data-ttu-id="88436-p117">Nom localis? de la police des en-t?tes. Sp?cifie le nom de la police par d?faut en fonction de la culture actuellement utilis?e dans Office.</span><span class="sxs-lookup"><span data-stu-id="88436-p117">Localized name of the headers font. Specifies the default font name according to the culture currently used in Office.</span></span>|

<br/>

#### <a name="theme-colors-for-content-add-ins"></a><span data-ttu-id="88436-189">Couleurs de th?me pour les compl?ments de contenu</span><span class="sxs-lookup"><span data-stu-id="88436-189">Theme colors for content add-ins</span></span>

|<span data-ttu-id="88436-190">**Classe**</span><span class="sxs-lookup"><span data-stu-id="88436-190">**Class**</span></span>|<span data-ttu-id="88436-191">**Description**</span><span class="sxs-lookup"><span data-stu-id="88436-191">**Description**</span></span>|
|:-----|:-----|
| `office-docTheme-primary-fontColor`|<span data-ttu-id="88436-p118">Couleur de police principale. Par d?faut : #000000</span><span class="sxs-lookup"><span data-stu-id="88436-p118">Primary font color. Default #000000</span></span>|
| `office-docTheme-primary-bgColor`|<span data-ttu-id="88436-p119">Couleur d?arri?re-plan de police principale. Par d?faut : #FFFFFF</span><span class="sxs-lookup"><span data-stu-id="88436-p119">Primary font background color. Default #FFFFFF</span></span>|
| `office-docTheme-secondary-fontColor`|<span data-ttu-id="88436-p120">Couleur de police secondaire. Par d?faut : #000000</span><span class="sxs-lookup"><span data-stu-id="88436-p120">Secondary font color. Default #000000</span></span>|
| `office-docTheme-secondary-bgColor`|<span data-ttu-id="88436-p121">Couleur d?arri?re-plan de police secondaire. Par d?faut : #FFFFFF</span><span class="sxs-lookup"><span data-stu-id="88436-p121">Secondary font background color. Default #FFFFFF</span></span>|
| `office-contentAccent1-color`|<span data-ttu-id="88436-p122">Couleur d?accentuation de police 1. Par d?faut : #5B9BD5</span><span class="sxs-lookup"><span data-stu-id="88436-p122">Font accent color 1. Default #5B9BD5</span></span>|
| `office-contentAccent2-color`|<span data-ttu-id="88436-p123">Couleur d?accentuation de police 2. Par d?faut : #ED7D31</span><span class="sxs-lookup"><span data-stu-id="88436-p123">Font accent color 2. Default #ED7D31</span></span>|
| `office-contentAccent3-color`|<span data-ttu-id="88436-p124">Couleur d?accentuation de police 3. Par d?faut : #A5A5A5</span><span class="sxs-lookup"><span data-stu-id="88436-p124">Font accent color 3. Default #A5A5A5</span></span>|
| `office-contentAccent4-color`|<span data-ttu-id="88436-p125">Couleur d?accentuation de police 4. Par d?faut : #FFC000</span><span class="sxs-lookup"><span data-stu-id="88436-p125">Font accent color 4. Default #FFC000</span></span>|
| `office-contentAccent5-color`|<span data-ttu-id="88436-p126">Couleur d?accentuation de police 5. Par d?faut : #4472C4</span><span class="sxs-lookup"><span data-stu-id="88436-p126">Font accent color 5. Default #4472C4</span></span>|
| `office-contentAccent6-color`|<span data-ttu-id="88436-p127">Couleur d?accentuation de police 6. Par d?faut : #70AD47</span><span class="sxs-lookup"><span data-stu-id="88436-p127">Font accent color 6. Default #70AD47</span></span>|
| `office-contentAccent1-bgColor`|<span data-ttu-id="88436-p128">Couleur d?accentuation d?arri?re-plan 1. Par d?faut : #5B9BD5</span><span class="sxs-lookup"><span data-stu-id="88436-p128">Background accent color 1. Default #5B9BD5</span></span>|
| `office-contentAccent2-bgColor`|<span data-ttu-id="88436-p129">Couleur d?accentuation d?arri?re-plan 2. Par d?faut : #ED7D31</span><span class="sxs-lookup"><span data-stu-id="88436-p129">Background accent color 2. Default #ED7D31</span></span>|
| `office-contentAccent3-bgColor`|<span data-ttu-id="88436-p130">Couleur d?accentuation d?arri?re-plan 3. Par d?faut : #A5A5A5</span><span class="sxs-lookup"><span data-stu-id="88436-p130">Background accent color 3. Default #A5A5A5</span></span>|
| `office-contentAccent4-bgColor`|<span data-ttu-id="88436-p131">Couleur d?accentuation d?arri?re-plan 4. Par d?faut : #FFC000</span><span class="sxs-lookup"><span data-stu-id="88436-p131">Background accent color 4. Default #FFC000</span></span>|
| `office-contentAccent5-bgColor`|<span data-ttu-id="88436-p132">Couleur d?accentuation d?arri?re-plan 5. Par d?faut : #4472C4</span><span class="sxs-lookup"><span data-stu-id="88436-p132">Background accent color 5. Default #4472C4</span></span>|
| `office-contentAccent6-bgColor`|<span data-ttu-id="88436-p133">Couleur d?accentuation d?arri?re-plan 6. Par d?faut : #70AD47</span><span class="sxs-lookup"><span data-stu-id="88436-p133">Background accent color 6. Default #70AD47</span></span>|
| `office-contentAccent1-borderColor`|<span data-ttu-id="88436-p134">Couleur d?accentuation de bordure 1. Par d?faut : #5B9BD5</span><span class="sxs-lookup"><span data-stu-id="88436-p134">Border accent color 1. Default #5B9BD5</span></span>|
| `office-contentAccent2-borderColor`|<span data-ttu-id="88436-p135">Couleur d?accentuation de bordure 2. Par d?faut : #ED7D31</span><span class="sxs-lookup"><span data-stu-id="88436-p135">Border accent color 2. Default #ED7D31</span></span>|
| `office-contentAccent3-borderColor`|<span data-ttu-id="88436-p136">Couleur d?accentuation de bordure 3. Par d?faut : #A5A5A5</span><span class="sxs-lookup"><span data-stu-id="88436-p136">Border accent color 3. Default #A5A5A5</span></span>|
| `office-contentAccent4-borderColor`|<span data-ttu-id="88436-p137">Couleur d?accentuation de bordure 4. Par d?faut : #FFC000</span><span class="sxs-lookup"><span data-stu-id="88436-p137">Border accent color 4. Default #FFC000</span></span>|
| `office-contentAccent5-borderColor`|<span data-ttu-id="88436-p138">Couleur d?accentuation de bordure 5. Par d?faut : #4472C4</span><span class="sxs-lookup"><span data-stu-id="88436-p138">Border accent color 5. Default #4472C4</span></span>|
| `office-contentAccent6-borderColor`|<span data-ttu-id="88436-p139">Couleur d?accentuation de bordure 6. Par d?faut : #70AD47</span><span class="sxs-lookup"><span data-stu-id="88436-p139">Border accent color 6. Default #70AD47</span></span>|
| `office-a`|<span data-ttu-id="88436-p140">Couleur de lien hypertexte. Par d?faut : #0563C1</span><span class="sxs-lookup"><span data-stu-id="88436-p140">Hyperlink color. Default #0563C1</span></span>|
| `office-a:visited`|<span data-ttu-id="88436-p141">Couleur de lien hypertexte visit?. Par d?faut : #954F72</span><span class="sxs-lookup"><span data-stu-id="88436-p141">Followed hyperlink color. Default #954F72</span></span>|

<br/>

<span data-ttu-id="88436-240">La capture d??cran suivante montre des exemples de toutes les classes de couleurs de th?me (sauf pour les deux couleurs de lien hypertexte) affect?es ? du texte d?compl?ment lorsque vous utilisez le th?me Office par d?faut.</span><span class="sxs-lookup"><span data-stu-id="88436-240">The following screenshot shows examples of all of the theme color classes (except for the two hyperlink colors) assigned to add-in text when using the default Office theme.</span></span>

![Exemple de couleurs de th?me Office par d?faut](../images/office15-app-default-office-theme-colors.png)


### <a name="theme-classes-for-task-pane-add-ins"></a><span data-ttu-id="88436-242">Classes de th?mes pour les compl?ments du volet Office</span><span class="sxs-lookup"><span data-stu-id="88436-242">Theme classes for task pane add-ins</span></span>

<span data-ttu-id="88436-p142">Le fichier OfficeThemes.css fournit des classes qui correspondent aux 4 couleurs affect?es aux polices et aux arri?re-plans utilis?s par le th?me de l?interface utilisateur de l?application Office. Ces classes peuvent ?tre utilis?es avec les compl?ments de t?che pour PowerPoint afin que les couleurs de votre compl?ment soient en harmonie avec les autres volets Office int?gr?s.</span><span class="sxs-lookup"><span data-stu-id="88436-p142">The OfficeThemes.css file provides classes that correspond to the 4 colors assigned to fonts and backgrounds used by the Office application UI theme. These classes are appropriate to use with task add-ins for PowerPoint so that your add-in's colors will be coordinated with the other built-in task panes in Office.</span></span>

#### <a name="theme-font-and-background-colors-for-task-pane-add-ins"></a><span data-ttu-id="88436-245">Couleurs de police et d?arri?re-plan de th?me pour les compl?ments du volet Office</span><span class="sxs-lookup"><span data-stu-id="88436-245">Theme font and background colors for task pane add-ins</span></span>

|<span data-ttu-id="88436-246">**Classe**</span><span class="sxs-lookup"><span data-stu-id="88436-246">**Class**</span></span>|<span data-ttu-id="88436-247">**Description**</span><span class="sxs-lookup"><span data-stu-id="88436-247">**Description**</span></span>|
|:-----|:-----|
| `office-officeTheme-primary-fontColor`|<span data-ttu-id="88436-p143">Couleur de police principale. Par d?faut : #B83B1D</span><span class="sxs-lookup"><span data-stu-id="88436-p143">Primary font color. Default #B83B1D</span></span>|
| `office-officeTheme-primary-bgColor`|<span data-ttu-id="88436-p144">Couleur d?arri?re-plan principale. Par d?faut : #DEDEDE</span><span class="sxs-lookup"><span data-stu-id="88436-p144">Primary background color. Default #DEDEDE</span></span>|
| `office-officeTheme-secondary-fontColor`|<span data-ttu-id="88436-p145">Couleur de police secondaire. Par d?faut : #262626.</span><span class="sxs-lookup"><span data-stu-id="88436-p145">Secondary font color. Default #262626</span></span>|
| `office-officeTheme-secondary-bgColor`|<span data-ttu-id="88436-p146">Couleur d?arri?re-plan secondaire. Par d?faut : #FFFFFF</span><span class="sxs-lookup"><span data-stu-id="88436-p146">Secondary background color. Default #FFFFFF</span></span>|

## <a name="see-also"></a><span data-ttu-id="88436-256">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="88436-256">See also</span></span>

- [<span data-ttu-id="88436-257">Cr?ation de compl?ments de contenu et du volet Office pour PowerPoint</span><span class="sxs-lookup"><span data-stu-id="88436-257">Create content and task pane add-ins for PowerPoint</span></span>](../powerpoint/powerpoint-add-ins.md)
