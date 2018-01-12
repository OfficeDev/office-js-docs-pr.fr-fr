# <a name="use-document-themes-in-your-powerpoint-add-ins"></a>Utiliser des thèmes de document dans vos compléments PowerPoint


Un [thème Office](https://support.office.com/en-US/Article/What-is-a-theme--7528ccc2-4327-4692-8bf5-9b5a3f2a5ef5) est constitué, en partie, d’un jeu de polices et de couleurs visuellement assortis que vous pouvez appliquer à des présentations, des documents, des feuilles de calcul et des courriers électroniques. Pour appliquer ou personnaliser le thème d’une présentation dans PowerPoint, utilisez les groupes **Thèmes** et **Variantes** dans l’onglet **Conception** du ruban. PowerPoint affecte le **thème Office** par défaut à chaque nouvelle présentation vierge, mais vous pouvez choisir parmi les autres thèmes disponibles dans l’onglet **Conception**, télécharger des thèmes supplémentaires à partir d’Office.com, ou créer et personnaliser votre propre thème.

OfficeThemes.css vous permet de concevoir des compléments coordonnés à PowerPoint de deux façons :


-  **Dans les compléments de contenu pour PowerPoint** Utilisez les classes de thèmes du document d’OfficeThemes.css pour spécifier les polices et les couleurs correspondant au thème de la présentation dans laquelle votre contenu complément est inséré ; ces polices et couleurs seront mises à jour dynamiquement si un utilisateur modifie ou personnalise le thème de la présentation.
    
-  **Dans les compléments du volet Office pour PowerPoint** Utilisez les classes de thèmes de l’interface utilisateur Office d’OfficeThemes.css pour spécifier les mêmes polices et couleurs d’arrière-plan que celles utilisées dans l’interface utilisateur, de sorte que vos compléments du volet Office correspondent aux couleurs des volets Office intégrés ; ces couleurs seront mises à jour dynamiquement si un utilisateur modifie le thème de l’interface utilisateur Office.
    

### <a name="document-theme-colors"></a>Couleurs de thème de document

Chaque thème de document Office définit 12 couleurs. Dix de ces couleurs sont disponibles lorsque vous définissez la police, l’arrière-plan et d’autres paramètres de couleur dans une présentation grâce au sélecteur de couleurs :


![Palette de couleurs](../../images/off15app_ColorPalette.png)

Pour afficher ou personnaliser l’intégralité des 12 couleurs de thème dans PowerPoint, dans le groupe **Variantes** de l’onglet **Conception**, cliquez sur le menu déroulant **Plus**, puis choisissez **Couleur** et cliquez sur **Personnaliser les couleurs** pour afficher la boîte de dialogue **Créer de nouvelles couleurs de thème** :


![Boîte de dialogue Créer de nouvelles couleurs de thème](../../images/off15app_CreateNewThemeColors.png)

Les quatre premières couleurs sont pour le texte et les arrière-plans. Un texte créé avec des couleurs claires sera toujours lisible sur les couleurs foncées, tandis qu’un texte créé avec des couleurs foncées sera toujours lisible sur les couleurs claires. Les six couleurs suivantes sont des couleurs d’accentuation qui sont toujours visibles sur les quatre couleurs d’arrière-plan potentielles. Les deux dernières couleurs sont pour les liens hypertexte et les liens hypertexte visités.


### <a name="document-theme-fonts"></a>Polices de thème de document

Chaque thème de document Office définit également deux polices : une pour les titres et l’autre pour le corps de texte. PowerPoint utilise ces polices pour créer des styles de texte automatiques. En outre, les galeries  **Styles rapides** pour le texte et **WordArt** utilisent ces mêmes polices de thème. Ces deux polices sont les deux premières proposées lorsque vous sélectionnez des polices avec le sélecteur de polices :


![Sélecteur de polices](../../images/off15app_FontPicker.png)

Pour afficher ou personnaliser les polices de thème dans PowerPoint, dans le groupe **Variantes** de l’onglet **Conception**, cliquez sur le menu déroulant **Plus**. Ensuite, pointez vers **Polices** et cliquez sur **Personnaliser les polices** pour afficher la boîte de dialogue **Créer de nouvelles polices de thème**.


![Boîte de dialogue Créer de nouvelles polices de thème](../../images/off15app_CreateNewThemeFonts.png)


### <a name="office-ui-theme-fonts-and-colors"></a>Couleurs et polices de thème de l’interface utilisateur Office

Office vous permet également de choisir entre plusieurs thèmes prédéfinis qui spécifient certaines des couleurs et des polices utilisées dans l’interface utilisateur de toutes les applications Office. Pour cela, utilisez le menu déroulant  **Fichier**  >   **Compte**  >   **Thème Office** (dans toutes les applications Office).


![Liste déroulante de thèmes Office](../../images/off15app_OfficeThemePicker.png)

OfficeThemes.css inclut des classes que vous pouvez utiliser dans vos compléments du volet Office pour PowerPoint afin qu’elles utilisent ces mêmes polices et couleurs. Cela vous permet de concevoir des compléments du volet Office dont l’apparence concorde avec celle des volets Office intégrés.


## <a name="using-officethemescss"></a>Utilisation d’OfficeThemes.css

En utilisant le fichier OfficeThemes.css avec vos compléments de contenu pour PowerPoint, vous pouvez coordonner l’apparence de votre complément avec le thème appliqué à la présentation avec laquelle elle est exécutée. En utilisant le fichier OfficeThemes.css avec vos compléments du volet Office pour PowerPoint, vous pouvez coordonner l’apparence de votre complément avec les polices et couleurs de l’interface utilisateur Office.


### <a name="adding-the-officethemescss-file-to-your-project"></a>Ajout du fichier OfficeThemes.css à votre projet

Suivez la procédure suivante pour ajouter et référencer le fichier OfficeThemes.css dans votre projet complément.


### <a name="to-add-officethemescss-to-your-visual-studio-project"></a>Pour ajouter le fichier OfficeThemes.css à votre projet Visual Studio


1. Dans l’ **Explorateur de solutions**, cliquez avec le bouton droit de la souris sur le dossier  **Contenu** dans le projet _**project_name**_**Web**, pointez vers **Ajouter** et cliquez sur **Feuille de style**.
    
2. Nommez la nouvelle feuille de style OfficeThemes.
    
     >**Important**  La feuille de style doit être nommée OfficeThemes, sinon la fonctionnalité qui met dynamiquement à jour les couleurs et les polices lorsqu’un utilisateur modifie le thème ne fonctionnera pas.
3. Supprimez la classe  **body** par défaut ( `body {}`) dans le fichier, et copiez-collez le code CSS suivant dans le fichier.
    
```
  /* The following classes describe the common theme information for office documents */ /* Basic Font and Background Colors for text */ .office-docTheme-primary-fontColor { color:#000000; } .office-docTheme-primary-bgColor { background-color:#ffffff; } .office-docTheme-secondary-fontColor { color: #000000; } .office-docTheme-secondary-bgColor { background-color: #ffffff; } /* Accent color definitions for fonts */ .office-contentAccent1-color { color:#5b9bd5; } .office-contentAccent2-color { color:#ed7d31; } .office-contentAccent3-color { color:#a5a5a5; } .office-contentAccent4-color { color:#ffc000; } .office-contentAccent5-color { color:#4472c4; } .office-contentAccent6-color { color:#70ad47; } /* Accent color for backgrounds */ .office-contentAccent1-bgColor { background-color:#5b9bd5; } .office-contentAccent2-bgColor { background-color:#ed7d31; } .office-contentAccent3-bgColor { background-color:#a5a5a5; } .office-contentAccent4-bgColor { background-color:#ffc000; } .office-contentAccent5-bgColor { background-color:#4472c4; } .office-contentAccent6-bgColor { background-color:#70ad47; } /* Accent color for borders */ .office-contentAccent1-borderColor { border-color:#5b9bd5; } .office-contentAccent2-borderColor { border-color:#ed7d31; } .office-contentAccent3-borderColor { border-color:#a5a5a5; } .office-contentAccent4-borderColor { border-color:#ffc000; } .office-contentAccent5-borderColor { border-color:#4472c4; } .office-contentAccent6-borderColor { border-color:#70ad47; } /* links */ .office-a { color: #0563c1; } .office-a:visited { color: #954f72; } /* Body Fonts */ .office-bodyFont-eastAsian { } /* East Asian name of the Font */ .office-bodyFont-latin { font-family:"Calibri"; } /* Latin name of the Font */ .office-bodyFont-script { } /* Script name of the Font */ .office-bodyFont-localized { font-family:"Calibri"; } /* Localized name of the Font. Corresponds to the default font of the culture currently used in Office.*/ /* Headers Font */ .office-headerFont-eastAsian { } .office-headerFont-latin { font-family:"Calibri Light"; } .office-headerFont-script { } .office-headerFont-localized { font-family:"Calibri Light"; } /* The following classes define font and background colors for Office UI themes. These classes should only be used in task pane add-ins */ /* Basic Font and Background Colors for PPT */ .office-officeTheme-primary-fontColor { color:#b83b1d; } .office-officeTheme-primary-bgColor { background-color:#dedede; } .office-officeTheme-secondary-fontColor { color:#262626; } .office-officeTheme-secondary-bgColor { background-color:#ffffff; } 
```

4. Si vous utilisez un autre outil que Visual Studio pour créer votre complément, copiez le code CSS de l’étape 3 dans un fichier texte, en vous assurant que le fichier est enregistré sous le nom OfficeThemes.css.
    

### <a name="referencing-officethemescss-in-your-add-ins-html-pages"></a>Référencement d’OfficeThemes.css dans les pages HTML de votre complément

Pour utiliser le fichier OfficeThemes.css dans votre projet complément, ajoutez une balise  `<link>` référençant le fichier OfficeThemes.css à l’intérieur de la balise `<head>` des pages web (par exemple, un fichier .html, .aspx ou .php) qui implémentent l’interface utilisateur de votre complément au format suivant :


```HTML
<link href="<local_path_to_OfficeThemes.css> " rel="stylesheet" type="text/css" />
```

Pour effectuer cette opération dans Visual Studio, procédez comme suit.


### <a name="to-reference-officethemescss-in-your-add-in-for-powerpoint"></a>Pour référencer OfficeThemes.css dans votre complément PowerPoint


1. Dans Visual Studio 2015, ouvrez ou créez un projet **Complément Office**.
    
2. Dans les pages HTML qui implémentent l’interface utilisateur de votre complément, telles que Home.html dans le modèle par défaut, ajoutez la balise  `<link>` suivante à l’intérieur de la balise `<head>` qui référence le fichier OfficeThemes.css :
    
```HTML
  <link href="../../Content/OfficeThemes.css" rel="stylesheet" type="text/css" />
```

Si vous créez votre complément avec un autre outil que Visual Studio, ajoutez une balise  `<link>` avec le même format spécifiant un chemin d’accès relatif vers la copie d’OfficeThemes.css qui sera déployée avec votre complément.


### <a name="using-officethemescss-document-theme-classes-in-your-content-add-ins-html-page"></a>Utilisation de classes de thèmes de document OfficeThemes.css dans la page HTML de votre complément de contenu

Ci-dessous figure un exemple simple de code HTML dans une complément de contenu qui utilise les classes de thèmes de document OfficeTheme.css. Pour plus d’informations sur les classes OfficeThemes.css qui correspondent aux 12 couleurs et aux 2 polices utilisées dans un thème de document, voir [Classes de thèmes pour les compléments de contenu](#theme-classes-for-content-add-ins).


```HTML
<body> <div id="themeSample" class="office-docTheme-primary-fontColor "> <h1 class="office-headerFont-latin">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent1-bgColor">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent2-bgColor">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent3-bgColor">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent4-bgColor">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent5-bgColor">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent6-bgColor">Hello world!</h1> <p class="office-bodyFont-latin office-docTheme-secondary-fontColor">Hello world!</p> </div> </body>
```

Lors de l’exécution, lorsque l’complément de contenu est insérée dans une présentation qui utilise le **thème Office** par défaut, elle est restituée comme suit :


![Application de contenu en cours d’exécution avec le thème Office](../../images/off15app_ContentApp_OfficeTheme.png)

Si vous modifiez la présentation afin d’utiliser un autre thème ou de personnaliser le thème de la présentation, les polices et couleurs spécifiées avec des classes OfficeThemes.css sont mises à jour dynamiquement pour correspondre aux polices et aux couleurs du thème de la présentation. En prenant l’exemple HTML ci-dessus, si la présentation dans laquelle le complément est inséré utilise le thème **Facette**, le complément est restitué comme suit :


![Application de contenu en cours d’exécution avec le thème Facette](../../images/off15app_ContentApp_FacetTheme.png)


### <a name="using-officethemescss-office-ui-theme-classes-in-your-task-pane-add-ins-html-page"></a>Utilisation de classes de thèmes de l’interface utilisateur Office OfficeThemes.css dans la page HTML de votre complément du volet Office

Outre le thème du document, les utilisateurs peuvent personnaliser le modèle de couleurs de l’interface utilisateur Office de toutes les applications Office à l’aide de la zone de liste déroulante  **Fichier**  >   **Compte**  >   **Thème Office**.

Ci-dessous figure un exemple simple de code HTML dans une complément de volet Office qui utilise des classes OfficeTheme.css pour spécifier les couleurs de police et d’arrière-plan. Pour plus d’informations sur les classes OfficeThemes.css qui correspondent aux polices et aux couleurs du thème de l’interface utilisateur Office, voir [Classes de thèmes pour les compléments du volet Office](#theme-classes-for-task-pane-add-ins).


```HTML
<body> <div id="content-header" class="office-officeTheme-primary-fontColor office-officeTheme-primary-bgColor"> <div class="padding"> <h1>Welcome</h1> </div> </div> <div id="content-main" class="office-officeTheme-secondary-fontColor office-officeTheme-secondary-bgColor"> <div class="padding"> <p>Add home screen content here.</p> <p>For example:</p> <button id="get-data-from-selection">Get data from selection</button> <p> <a target="_blank" class="office-a" href="https://go.microsoft.com/fwlink/?LinkId=276812"> Find more samples online... </a> </p> </div> </div> </body> 
```

Lors de l’exécution de PowerPoint avec  **Fichier**  >   **Compte**  >   **Thème Office** défini sur **Blanc**, l’complémentde volet de tâches est restituée comme suit :


![Volet de tâches avec thème blanc Office](../../images/off15app_TaskPaneThemeWhite.png)

Si vous modifiez la valeur de **Thème Office** en la définissant sur **Gris foncé**, les polices et couleurs spécifiées avec des classes OfficeThemes.css seront mises à jour dynamiquement et seront restituées comme suit :


![Volet de tâches avec thème gris foncé Office](../../images/off15app_TaskPaneThemeDarkGray.png)


## <a name="officethemecss-classes"></a>Classes OfficeTheme.css


Le fichier OfficeThemes.css contient deux jeux de classes que vous pouvez utiliser avec vos compléments de contenu et du volet Office PowerPoint.


### <a name="theme-classes-for-content-add-ins"></a>Classes de thèmes pour les compléments de contenu


Le fichier OfficeThemes.css fournit des classes qui correspondent aux 12 couleurs et aux 2 polices utilisées dans un thème de document. Ces classes sont adaptées aux compléments de contenu pour PowerPoint, de sorte que les polices et les couleurs de votre complément seront en harmonie avec la présentation dans laquelle votre complément est inséré.


**Polices de thème pour les compléments de contenu**


|**Classe**|**Description**|
|:-----|:-----|
| `office-bodyFont-eastAsian`|Nom en langues d’Asie de l’Est de la police du corps de texte.|
| `office-bodyFont-latin`|Nom latin de la police du corps de texte (par défaut, « Calibri »).|
| `office-bodyFont-script`|Nom de script de la police du corps de texte.|
| `office-bodyFont-localized`|Nom localisé de la police du corps de texte. Spécifie le nom de la police par défaut en fonction de la culture actuellement utilisée dans Office.|
| `office-headerFont-eastAsian`|Nom en langues d’Asie de l’Est de la police des en-têtes.|
| `office-headerFont-latin`|Nom latin de la police des en-têtes (par défaut, « Calibri Light »).|
| `office-headerFont-script`|Nom de script de la police des en-têtes.|
| `office-headerFont-localized`|Nom localisé de la police des en-têtes. Spécifie le nom de la police par défaut en fonction de la culture actuellement utilisée dans Office.|

**Couleurs de thème pour les compléments de contenu**


|**Classe**|**Description**|
|:-----|:-----|
| `office-docTheme-primary-fontColor`|Couleur de police principale. Par défaut : #000000|
| `office-docTheme-primary-bgColor`|Couleur d’arrière-plan de police principale. Par défaut : #FFFFFF|
| `office-docTheme-secondary-fontColor`|Couleur de police secondaire. Par défaut : #000000|
| `office-docTheme-secondary-bgColor`|Couleur d’arrière-plan de police secondaire. Par défaut : #FFFFFF|
| `office-contentAccent1-color`|Couleur d’accentuation de police 1. Par défaut : #5B9BD5|
| `office-contentAccent2-color`|Couleur d’accentuation de police 2. Par défaut : #ED7D31|
| `office-contentAccent3-color`|Couleur d’accentuation de police 3. Par défaut : #A5A5A5|
| `office-contentAccent4-color`|Couleur d’accentuation de police 4. Par défaut : #FFC000|
| `office-contentAccent5-color`|Couleur d’accentuation de police 5. Par défaut : #4472C4|
| `office-contentAccent6-color`|Couleur d’accentuation de police 6. Par défaut : #70AD47|
| `office-contentAccent1-bgColor`|Couleur d’accentuation d’arrière-plan 1. Par défaut : #5B9BD5|
| `office-contentAccent2-bgColor`|Couleur d’accentuation d’arrière-plan 2. Par défaut : #ED7D31|
| `office-contentAccent3-bgColor`|Couleur d’accentuation d’arrière-plan 3. Par défaut : #A5A5A5|
| `office-contentAccent4-bgColor`|Couleur d’accentuation d’arrière-plan 4. Par défaut : #FFC000|
| `office-contentAccent5-bgColor`|Couleur d’accentuation d’arrière-plan 5. Par défaut : #4472C4|
| `office-contentAccent6-bgColor`|Couleur d’accentuation d’arrière-plan 6. Par défaut : #70AD47|
| `office-contentAccent1-borderColor`|Couleur d’accentuation de bordure 1. Par défaut : #5B9BD5|
| `office-contentAccent2-borderColor`|Couleur d’accentuation de bordure 2. Par défaut : #ED7D31|
| `office-contentAccent3-borderColor`|Couleur d’accentuation de bordure 3. Par défaut : #A5A5A5|
| `office-contentAccent4-borderColor`|Couleur d’accentuation de bordure 4. Par défaut : #FFC000|
| `office-contentAccent5-borderColor`|Couleur d’accentuation de bordure 5. Par défaut : #4472C4|
| `office-contentAccent6-borderColor`|Couleur d’accentuation de bordure 6. Par défaut : #70AD47|
| `office-a`|Couleur de lien hypertexte. Par défaut : #0563C1|
| `office-a:visited`|Couleur de lien hypertexte visité. Par défaut : #954F72|
La capture d’écran suivante montre des exemples de toutes les classes de couleurs de thème (sauf pour les deux couleurs de lien hypertexte) affectées à du texte d’complément lorsque vous utilisez le thème Office par défaut.


![Exemple de couleurs de thème Office par défaut](../../images/off15app_DefaultOfficeThemeColors.png)


### <a name="theme-classes-for-task-pane-add-ins"></a>Classes de thèmes pour les compléments du volet Office


Le fichier OfficeThemes.css fournit des classes qui correspondent aux 4 couleurs affectées aux polices et aux arrière-plans utilisés par le thème de l’interface utilisateur de l’application Office. Ces classes peuvent être utilisées avec les compléments de tâche pour PowerPoint afin que les couleurs de votre complément soient en harmonie avec les autres volets Office intégrés.


**Couleurs de police et d’arrière-plan de thème pour les compléments du volet Office**


|**Classe**|**Description**|
|:-----|:-----|
| `office-officeTheme-primary-fontColor`|Couleur de police principale. Par défaut : #B83B1D|
| `office-officeTheme-primary-bgColor`|Couleur d’arrière-plan principale. Par défaut : #DEDEDE|
| `office-officeTheme-secondary-fontColor`|Couleur de police secondaire. Par défaut : #262626.|
| `office-officeTheme-secondary-bgColor`|Couleur d’arrière-plan secondaire. Par défaut : #FFFFFF|

## <a name="additional-resources"></a>Ressources supplémentaires

- [Création de compléments de contenu et du volet Office pour PowerPoint](../powerpoint/powerpoint-add-ins.md)
