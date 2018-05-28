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
# <a name="use-document-themes-in-your-powerpoint-add-ins"></a>Utiliser des th?mes de document dans vos compl?ments PowerPoint

Un [th?me Office](https://support.office.com/en-US/Article/What-is-a-theme--7528ccc2-4327-4692-8bf5-9b5a3f2a5ef5) est constitu?, en partie, d?un jeu de polices et de couleurs visuellement assortis que vous pouvez appliquer ? des pr?sentations, des documents, des feuilles de calcul et des courriers ?lectroniques. Pour appliquer ou personnaliser le th?me d?une pr?sentation dans PowerPoint, utilisez les groupes **Th?mes** et **Variantes** dans l?onglet **Conception** du ruban. PowerPoint affecte le **th?me Office** par d?faut ? chaque nouvelle pr?sentation vierge, mais vous pouvez choisir parmi les autres th?mes disponibles dans l?onglet **Conception**, t?l?charger des th?mes suppl?mentaires ? partir d?Office.com, ou cr?er et personnaliser votre propre th?me.

OfficeThemes.css vous permet de concevoir des compl?ments coordonn?s ? PowerPoint de deux fa?ons :

- **Dans les compl?ments de contenu pour PowerPoint**. Utilisez les classes de th?mes du document d?OfficeThemes.css pour sp?cifier les polices et les couleurs correspondant au th?me de la pr?sentation dans laquelle votre contenu compl?ment est ins?r? ; ces polices et couleurs seront mises ? jour dynamiquement si un utilisateur modifie ou personnalise le th?me de la pr?sentation.
    
- **Dans les compl?ments du volet Office pour PowerPoint**. Utilisez les classes de th?mes de l?interface utilisateur Office d?OfficeThemes.css pour sp?cifier les m?mes polices et couleurs d?arri?re-plan que celles utilis?es dans l?interface utilisateur, de sorte que vos compl?ments du volet Office correspondent aux couleurs des volets Office int?gr?s ; ces couleurs seront mises ? jour dynamiquement si un utilisateur modifie le th?me de l?interface utilisateur Office.

### <a name="document-theme-colors"></a>Couleurs de th?me de document

Chaque th?me de document Office d?finit 12 couleurs. Dix de ces couleurs sont disponibles lorsque vous d?finissez la police, l?arri?re-plan et d?autres param?tres de couleur dans une pr?sentation gr?ce au s?lecteur de couleurs.

![Palette de couleurs](../images/office15-app-color-palette.png)

Pour afficher ou personnaliser l?int?gralit? des 12 couleurs de th?me dans PowerPoint, dans le groupe **Variantes** de l?onglet **Conception**, cliquez sur le menu d?roulant **Plus**, puis choisissez **Couleur** et cliquez sur **Personnaliser les couleurs** pour afficher la bo?te de dialogue **Cr?er de nouvelles couleurs de th?me**.

![Bo?te de dialogue Cr?er de nouvelles couleurs de th?me](../images/office15-app-create-new-theme-colors.png)

Les quatre premi?res couleurs sont pour le texte et les arri?re-plans. Un texte cr?? avec des couleurs claires sera toujours lisible sur les couleurs fonc?es, tandis qu?un texte cr?? avec des couleurs fonc?es sera toujours lisible sur les couleurs claires. Les six couleurs suivantes sont des couleurs d?accentuation qui sont toujours visibles sur les quatre couleurs d?arri?re-plan potentielles. Les deux derni?res couleurs sont pour les liens hypertexte et les liens hypertexte visit?s.

### <a name="document-theme-fonts"></a>Polices de th?me de document

Chaque th?me de document Office d?finit ?galement deux polices : une pour les titres et l?autre pour le corps de texte. PowerPoint utilise ces polices pour cr?er des styles de texte automatiques. En outre, les galeries **Styles rapides** pour le texte et **WordArt** utilisent ces m?mes polices de th?me. Ces deux polices sont les deux premi?res propos?es lorsque vous s?lectionnez des polices avec le s?lecteur de polices.

![S?lecteur de polices](../images/office15-app-font-picker.png)

Pour afficher ou personnaliser les polices de th?me dans PowerPoint, dans le groupe **Variantes** de l?onglet **Conception**, cliquez sur le menu d?roulant **Plus**. Ensuite, pointez vers **Polices** et cliquez sur **Personnaliser les polices** pour afficher la bo?te de dialogue **Cr?er de nouvelles polices de th?me**.

![Bo?te de dialogue Cr?er de nouvelles polices de th?me](../images/office15-app-create-new-theme-fonts.png)

### <a name="office-ui-theme-fonts-and-colors"></a>Couleurs et polices de th?me de l?interface utilisateur Office

Office vous permet ?galement de choisir entre plusieurs th?mes pr?d?finis qui sp?cifient certaines des couleurs et des polices utilis?es dans l?interface utilisateur de toutes les applications Office. Pour cela, utilisez le menu d?roulant  **Fichier**  >   **Compte**  >   **Th?me Office** (dans toutes les applications Office).

![Liste d?roulante de th?mes Office](../images/office15-app-office-theme-picker.png)

OfficeThemes.css inclut des classes que vous pouvez utiliser dans vos compl?ments du volet Office pour PowerPoint afin qu?elles utilisent ces m?mes polices et couleurs. Cela vous permet de concevoir des compl?ments du volet Office dont l?apparence concorde avec celle des volets Office int?gr?s.

## <a name="using-officethemescss"></a>Utilisation d?OfficeThemes.css

En utilisant le fichier OfficeThemes.css avec vos compl?ments de contenu pour PowerPoint, vous pouvez coordonner l?apparence de votre compl?ment avec le th?me appliqu? ? la pr?sentation avec laquelle elle est ex?cut?e. En utilisant le fichier OfficeThemes.css avec vos compl?ments du volet Office pour PowerPoint, vous pouvez coordonner l?apparence de votre compl?ment avec les polices et couleurs de l?interface utilisateur Office.

### <a name="adding-the-officethemescss-file-to-your-project"></a>Ajout du fichier OfficeThemes.css ? votre projet

Suivez la proc?dure suivante pour ajouter et r?f?rencer le fichier OfficeThemes.css dans votre projet compl?ment.

#### <a name="to-add-officethemescss-to-your-visual-studio-project"></a>Pour ajouter le fichier OfficeThemes.css ? votre projet Visual Studio

1. Dans l?**explorateur de solutions**, cliquez avec le bouton droit de la souris sur le dossier **Contenu** dans le projet _**project_name**_**Web**, pointez sur **Ajouter** et s?lectionnez **Feuille de style**.
    
2. Nommez la nouvelle feuille de style **OfficeThemes**.
    
   > [!IMPORTANT]
   > Le nom de la feuille de style doit ?tre OfficeThemes, sinon la fonctionnalit? qui met ? jour dynamiquement les polices et couleurs de compl?ment lorsqu?un utilisateur modifie le th?me ne fonctionnera pas.
   
3. Supprimez la classe **body** par d?faut (`body {}`) dans le fichier, et copiez-collez le code CSS suivant dans le fichier.
    
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
4. Si vous utilisez un autre outil que Visual Studio pour cr?er votre compl?ment, copiez le code CSS de l??tape 3 dans un fichier texte, en vous assurant que le fichier est enregistr? sous le nom OfficeThemes.css.   

### <a name="referencing-officethemescss-in-your-add-ins-html-pages"></a>R?f?rencement d?OfficeThemes.css dans les pages HTML de votre compl?ment

Pour utiliser le fichier OfficeThemes.css dans votre projet de compl?ment, ajoutez une balise `<link>` r?f?ren?ant le fichier OfficeThemes.css ? l?int?rieur de la balise `<head>` des pages web (par exemple, un fichier .html, .aspx ou .php) qui impl?mentent l?interface utilisateur de votre compl?ment au format suivant :

```HTML
<link href="<local_path_to_OfficeThemes.css>" rel="stylesheet" type="text/css" />
```

Pour effectuer cette op?ration dans Visual Studio, proc?dez comme suit.

#### <a name="to-reference-officethemescss-in-your-add-in-for-powerpoint"></a>Pour r?f?rencer OfficeThemes.css dans votre compl?ment PowerPoint

1. Dans Visual Studio 2015, ouvrez ou cr?ez un projet de **compl?ment Office**.
    
2. Dans les pages HTML qui impl?mentent l?interface utilisateur de votre compl?ment, telles que Home.html dans le mod?le par d?faut, ajoutez la balise `<link>` suivante ? l?int?rieur de la balise `<head>` qui r?f?rence le fichier OfficeThemes.css :
    
    ```HTML
    <link href="../../Content/OfficeThemes.css" rel="stylesheet" type="text/css" />
    ```

Si vous cr?ez votre compl?ment avec un outil autre que Visual Studio, ajoutez une balise `<link>` avec le m?me format sp?cifiant un chemin d?acc?s relatif vers la copie d?OfficeThemes.css qui sera d?ploy?e avec votre compl?ment.

### <a name="using-officethemescss-document-theme-classes-in-your-content-add-ins-html-page"></a>Utilisation de classes de th?mes de document OfficeThemes.css dans la page HTML de votre compl?ment de contenu

Ci-dessous figure un exemple simple de code HTML dans une compl?ment de contenu qui utilise les classes de th?mes de document OfficeTheme.css. Pour plus d?informations sur les classes OfficeThemes.css qui correspondent aux 12 couleurs et aux 2 polices utilis?es dans un th?me de document, voir [Classes de th?mes pour les compl?ments de contenu](#theme-classes-for-content-add-ins).

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

Lors de l?ex?cution, lorsque le compl?ment de contenu est ins?r? dans une pr?sentation qui utilise le **th?me Office** par d?faut, il est restitu? comme suit.

![Application de contenu en cours d?ex?cution avec le th?me Office](../images/office15-app-content-app-office-theme.png)

Si vous modifiez la pr?sentation afin d?utiliser un autre th?me ou de personnaliser le th?me de la pr?sentation, les polices et couleurs sp?cifi?es avec des classes OfficeThemes.css sont mises ? jour dynamiquement pour correspondre aux polices et aux couleurs du th?me de la pr?sentation. En prenant l?exemple HTML ci-dessus, si la pr?sentation dans laquelle le compl?ment est ins?r? utilise le th?me **Facette**, le compl?ment est restitu? comme suit.

![Application de contenu en cours d?ex?cution avec le th?me Facette](../images/office15-app-content-app-facet-theme.png)


### <a name="using-officethemescss-office-ui-theme-classes-in-your-task-pane-add-ins-html-page"></a>Utilisation de classes de th?mes de l?interface utilisateur Office OfficeThemes.css dans la page HTML de votre compl?ment du volet Office

Outre le th?me du document, les utilisateurs peuvent personnaliser le mod?le de couleurs de l?interface utilisateur Office de toutes les applications Office ? l?aide de la zone de liste d?roulante **Fichier** > **Compte** > **Th?me Office**.

Ci-dessous figure un exemple simple de code HTML dans une compl?ment de volet Office qui utilise des classes OfficeTheme.css pour sp?cifier les couleurs de police et d?arri?re-plan. Pour plus d?informations sur les classes OfficeThemes.css qui correspondent aux polices et aux couleurs du th?me de l?interface utilisateur Office, voir [Classes de th?mes pour les compl?ments du volet Office](#theme-classes-for-task-pane-add-ins).

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

Lors de l?ex?cution de PowerPoint avec **Fichier** > **Compte** > **Th?me Office** d?fini sur **Blanc**, le compl?ment de volet de t?ches est restitu? comme suit.

![Volet de t?ches avec th?me blanc Office](../images/office15-app-task-pane-theme-white.png)

<br/>

Si vous modifiez la valeur de **Th?me Office** en la d?finissant sur **Gris fonc?**, les polices et couleurs sp?cifi?es avec des classes OfficeThemes.css seront mises ? jour dynamiquement et seront restitu?es comme suit.

![Volet de t?ches avec th?me gris fonc? Office](../images/office15-app-task-pane-theme-dark-gray.png)

<br/>

## <a name="officethemecss-classes"></a>Classes OfficeTheme.css

Le fichier OfficeThemes.css contient deux jeux de classes que vous pouvez utiliser avec vos compl?ments de contenu et du volet Office PowerPoint.

### <a name="theme-classes-for-content-add-ins"></a>Classes de th?mes pour les compl?ments de contenu

Le fichier OfficeThemes.css fournit des classes qui correspondent aux 12 couleurs et aux 2 polices utilis?es dans un th?me de document. Ces classes sont adapt?es aux compl?ments de contenu pour PowerPoint, de sorte que les polices et les couleurs de votre compl?ment seront en harmonie avec la pr?sentation dans laquelle votre compl?ment est ins?r?.

#### <a name="theme-fonts-for-content-add-ins"></a>Polices de th?me pour les compl?ments de contenu

|**Classe**|**Description**|
|:-----|:-----|
| `office-bodyFont-eastAsian`|Nom en langues d?Asie de l?Est de la police du corps de texte.|
| `office-bodyFont-latin`|Nom latin de la police du corps de texte (par d?faut, ? Calibri ?).|
| `office-bodyFont-script`|Nom de script de la police du corps de texte.|
| `office-bodyFont-localized`|Nom localis? de la police du corps de texte. Sp?cifie le nom de la police par d?faut en fonction de la culture actuellement utilis?e dans Office.|
| `office-headerFont-eastAsian`|Nom en langues d?Asie de l?Est de la police des en-t?tes.|
| `office-headerFont-latin`|Nom latin de la police des en-t?tes (par d?faut, ? Calibri Light ?).|
| `office-headerFont-script`|Nom de script de la police des en-t?tes.|
| `office-headerFont-localized`|Nom localis? de la police des en-t?tes. Sp?cifie le nom de la police par d?faut en fonction de la culture actuellement utilis?e dans Office.|

<br/>

#### <a name="theme-colors-for-content-add-ins"></a>Couleurs de th?me pour les compl?ments de contenu

|**Classe**|**Description**|
|:-----|:-----|
| `office-docTheme-primary-fontColor`|Couleur de police principale. Par d?faut : #000000|
| `office-docTheme-primary-bgColor`|Couleur d?arri?re-plan de police principale. Par d?faut : #FFFFFF|
| `office-docTheme-secondary-fontColor`|Couleur de police secondaire. Par d?faut : #000000|
| `office-docTheme-secondary-bgColor`|Couleur d?arri?re-plan de police secondaire. Par d?faut : #FFFFFF|
| `office-contentAccent1-color`|Couleur d?accentuation de police 1. Par d?faut : #5B9BD5|
| `office-contentAccent2-color`|Couleur d?accentuation de police 2. Par d?faut : #ED7D31|
| `office-contentAccent3-color`|Couleur d?accentuation de police 3. Par d?faut : #A5A5A5|
| `office-contentAccent4-color`|Couleur d?accentuation de police 4. Par d?faut : #FFC000|
| `office-contentAccent5-color`|Couleur d?accentuation de police 5. Par d?faut : #4472C4|
| `office-contentAccent6-color`|Couleur d?accentuation de police 6. Par d?faut : #70AD47|
| `office-contentAccent1-bgColor`|Couleur d?accentuation d?arri?re-plan 1. Par d?faut : #5B9BD5|
| `office-contentAccent2-bgColor`|Couleur d?accentuation d?arri?re-plan 2. Par d?faut : #ED7D31|
| `office-contentAccent3-bgColor`|Couleur d?accentuation d?arri?re-plan 3. Par d?faut : #A5A5A5|
| `office-contentAccent4-bgColor`|Couleur d?accentuation d?arri?re-plan 4. Par d?faut : #FFC000|
| `office-contentAccent5-bgColor`|Couleur d?accentuation d?arri?re-plan 5. Par d?faut : #4472C4|
| `office-contentAccent6-bgColor`|Couleur d?accentuation d?arri?re-plan 6. Par d?faut : #70AD47|
| `office-contentAccent1-borderColor`|Couleur d?accentuation de bordure 1. Par d?faut : #5B9BD5|
| `office-contentAccent2-borderColor`|Couleur d?accentuation de bordure 2. Par d?faut : #ED7D31|
| `office-contentAccent3-borderColor`|Couleur d?accentuation de bordure 3. Par d?faut : #A5A5A5|
| `office-contentAccent4-borderColor`|Couleur d?accentuation de bordure 4. Par d?faut : #FFC000|
| `office-contentAccent5-borderColor`|Couleur d?accentuation de bordure 5. Par d?faut : #4472C4|
| `office-contentAccent6-borderColor`|Couleur d?accentuation de bordure 6. Par d?faut : #70AD47|
| `office-a`|Couleur de lien hypertexte. Par d?faut : #0563C1|
| `office-a:visited`|Couleur de lien hypertexte visit?. Par d?faut : #954F72|

<br/>

La capture d??cran suivante montre des exemples de toutes les classes de couleurs de th?me (sauf pour les deux couleurs de lien hypertexte) affect?es ? du texte d?compl?ment lorsque vous utilisez le th?me Office par d?faut.

![Exemple de couleurs de th?me Office par d?faut](../images/office15-app-default-office-theme-colors.png)


### <a name="theme-classes-for-task-pane-add-ins"></a>Classes de th?mes pour les compl?ments du volet Office

Le fichier OfficeThemes.css fournit des classes qui correspondent aux 4 couleurs affect?es aux polices et aux arri?re-plans utilis?s par le th?me de l?interface utilisateur de l?application Office. Ces classes peuvent ?tre utilis?es avec les compl?ments de t?che pour PowerPoint afin que les couleurs de votre compl?ment soient en harmonie avec les autres volets Office int?gr?s.

#### <a name="theme-font-and-background-colors-for-task-pane-add-ins"></a>Couleurs de police et d?arri?re-plan de th?me pour les compl?ments du volet Office

|**Classe**|**Description**|
|:-----|:-----|
| `office-officeTheme-primary-fontColor`|Couleur de police principale. Par d?faut : #B83B1D|
| `office-officeTheme-primary-bgColor`|Couleur d?arri?re-plan principale. Par d?faut : #DEDEDE|
| `office-officeTheme-secondary-fontColor`|Couleur de police secondaire. Par d?faut : #262626.|
| `office-officeTheme-secondary-bgColor`|Couleur d?arri?re-plan secondaire. Par d?faut : #FFFFFF|

## <a name="see-also"></a>Voir aussi

- [Cr?ation de compl?ments de contenu et du volet Office pour PowerPoint](../powerpoint/powerpoint-add-ins.md)
