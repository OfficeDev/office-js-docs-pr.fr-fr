---
title: Localisation des compl?ments Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: d7888859ca29a62541020b45b0b7a3638c41f4f2
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="localization-for-office-add-ins"></a>Localisation des compl?ments Office

Vous pouvez librement impl?menter n?importe quel sch?ma de localisation convenant ? votre Compl?ment Office. L?API JavaScript et le sch?ma du manifeste de la plateforme Compl?ments Office offrent quelques choix. Vous pouvez utiliser l?API JavaScript pour Office pour d?terminer un param?tre r?gional et les cha?nes d?affichage en fonction des param?tres r?gionaux de l?application h?te, ou pour interpr?ter ou afficher les donn?es en fonction des param?tres r?gionaux des donn?es. Vous pouvez utiliser le manifeste pour sp?cifier l?emplacement des fichiers et les informations descriptives propres ? un param?tre r?gional. Sinon, vous pouvez utiliser un script Microsoft Ajax pour prendre en charge l?internationalisation et la localisation.

## <a name="use-the-javascript-api-to-determine-locale-specific-strings"></a>Utiliser l?API JavaScript pour d?terminer les cha?nes propres aux param?tres r?gionaux

L?API JavaScript pour Office offre deux propri?t?s qui prennent en charge l?affichage ou l?interpr?tation de valeurs coh?rentes avec les param?tres r?gionaux de l?application h?te et des donn?es :

- [Context.displayLanguage][displayLanguage] sp?cifie les param?tres r?gionaux (ou langue) de l?interface utilisateur de l?application h?te. L?exemple suivant v?rifie si l?application h?te utilise les param?tres r?gionaux en-US ou fr-Fr, et affiche un message de bienvenue propre aux param?tres r?gionaux.
    
    ```js
    function sayHelloWithDisplayLanguage() {
        var myLanguage = Office.context.displayLanguage;
        switch (myLanguage) {
            case 'en-US':
                write('Hello!');
                break;
            case 'fr-FR':
                write('Bonjour!');
                break;
        }
    }
    
    // Function that writes to a div with id='message' on the page.
    function write(message) {
        document.getElementById('message').innerText += message; 
    }
    ```

- [Context.contentLanguage][contentLanguage] sp?cifie le param?tre r?gional (ou langue) des donn?es. Le fait d??tendre le dernier exemple de code, au lieu de v?rifier la propri?t? [displayLanguage], attribue `myLanguage` ? la propri?t? [contentLanguage] et utilise le reste du code pour afficher un message de bienvenue correspondant aux param?tres r?gionaux des donn?es :
    
    ```js
    var myLanguage = Office.context.contentLanguage;
    ```

## <a name="control-localization-from-the-manifest"></a>Contr?ler la localisation ? partir du manifeste


Chaque compl?ment Office indique un ?l?ment [DefaultLocale] ?l?ment et un param?tre r?gional dans son manifeste. Par d?faut, la plateforme de compl?ment Office et les applications h?tes Office appliquent les valeurs des ?l?ments [Description], [DisplayName], [IconUrl], [HighResolutionIconUrl] et [SourceLocation] ? tous les param?tres r?gionaux. Vous pouvez ?ventuellement prendre en charge des valeurs sp?cifiques pour les param?tres r?gionaux sp?cifiques, en sp?cifiant un ?l?ment enfant [Override] pour chaque param?tre r?gional suppl?mentaire, pour chacun des cinq ?l?ments. La valeur de l??l?ment [DefaultLocale] et de l?attribut `Locale` de l??l?ment [Override] est sp?cifi?e en fonction de la norme [RFC 3066] relative aux balises pour l?identification des langues (? Tags for the Identification of Languages ?). Le tableau 1 d?crit la prise en charge de localisation de ces ?l?ments.

**Tableau 1. Prise en charge de localisation**


|**?l?ment**|**Prise en charge de localisation**|
|:-----|:-----|
|[Description]   |Les utilisateurs de chaque param?tre r?gional que vous sp?cifiez peuvent voir une description localis?e du compl?ment dans AppSource (ou dans un catalogue priv?).<br/>Pour les compl?ments Outlook, les utilisateurs peuvent voir la description dans le Centre d?administration Exchange (EAC) apr?s l?installation.|
|[DisplayName]   |Les utilisateurs de chaque param?tre r?gional que vous sp?cifiez peuvent voir une description localis?e du compl?ment dans AppSource (ou dans un catalogue priv?).<br/>Pour les compl?ments Outlook, les utilisateurs peuvent voir le nom d?affichage sous forme d??tiquette pour le bouton de l?application Outlook ainsi que dans l?EAC apr?s l?installation.<br/>Pour les compl?ments de contenu et du volet Office, les utilisateurs peuvent voir l?ic?ne dans le ruban apr?s avoir install? l?application.|
|[IconUrl]        |L?image de l?ic?ne est facultative. Vous pouvez utiliser la m?me technique de remplacement pour sp?cifier une image donn?e pour une culture particuli?re. Si vous utilisez et localisez une ic?ne, les utilisateurs de chaque param?tre r?gional que vous sp?cifiez peuvent voir l?image d?ic?ne localis?e pour le compl?ment.<br/>Pour les compl?ments Outlook, les utilisateurs peuvent voir l?ic?ne dans l?EAC apr?s l?installation du compl?ment.<br/>Pour les compl?ments de contenu et du volet de t?ches, les utilisateurs peuvent voir l?ic?ne dans le ruban apr?s avoir install? le compl?ment.|
|[HighResolutionIconUrl] **Important :** cet ?l?ment est disponible uniquement lors de l?utilisation de la version 1.1 du manifeste de compl?ment.|L?image de l?ic?ne de haute r?solution est facultative. N?anmoins, si elle est indiqu?e, elle doit l??tre apr?s l??l?ment [IconUrl]. Si  [HighResolutionIconUrl] est sp?cifi? et que le compl?ment est install? sur un appareil qui prend en charge la haute r?solution (dpi), la valeur [HighResolutionIconUrl] est utilis?e ? la place de la valeur [IconUrl].<br/>Vous pouvez utiliser la m?me technique de remplacement pour sp?cifier une image donn?e pour une culture particuli?re. Si vous utilisez et localisez une ic?ne, les utilisateurs de chaque param?tre r?gional que vous sp?cifiez peuvent voir l?image d?ic?ne localis?e pour le compl?ment.<br/>Pour les compl?ments Outlook, les utilisateurs peuvent voir l?ic?ne dans l?EAC apr?s l?installation du compl?ment.<br/>Pour les compl?ments de contenu et du volet de t?ches, les utilisateurs peuvent voir l?ic?ne dans le ruban apr?s avoir install? le compl?ment.|
|[Ressources] **Important :** cet ?l?ment est disponible uniquement lors de l?utilisation de la version 1.1 du manifeste de compl?ment.   |Les utilisateurs de chaque param?tre r?gional que vous sp?cifiez peuvent voir les ressources de cha?ne et d?ic?ne que vous cr?ez sp?cifiquement pour le compl?ment pour ce param?tre r?gional. |
|[SourceLocation]   |Les utilisateurs de chaque param?tre r?gional que vous sp?cifiez peuvent voir une page web que vous concevez sp?cifiquement pour le compl?ment pour ce param?tre r?gional. |


> **REMARQUE** Vous pouvez trouver la description et le nom d?affichage uniquement pour les param?tres r?gionaux pris en charge par Office. Reportez-vous ? la rubrique [Identificateurs de langue et valeurs d'ID de l'?l?ment OptionState dans Office 2013](http://technet.microsoft.com/en-us/library/cc179219.aspx) pour conna?tre la liste des langues et des param?tres r?gionaux pour la version actuelle d?Office.


### <a name="examples"></a>Exemples

Par exemple, un compl?ment Office peut sp?cifier [DefaultLocale] en tant que `en-us`. Pour l??l?ment [DisplayName], le compl?ment peut sp?cifier un ?l?ment enfant [Override] pour le param?tre r?gional `fr-fr`, comme illustr? ci-dessous. 


```xml
<DefaultLocale>en-us</DefaultLocale>
...
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

> **REMARQUE** Si vous devez rechercher plusieurs domaines au sein d?une famille de langues, comme `de-de` et `de-at`, nous vous recommandons d?utiliser des ?l?ments `Override` distincts pour chaque domaine. L?utilisation uniquement du nom de la langue, soit `de` dans ce cas, n?est pas prise en charge pour toutes les combinaisons de plateformes et d?applications h?te Office.

Cela signifie que le compl?ment adopte le param?tre r?gional `en-us` par d?faut. Les utilisateurs voient le nom d?affichage ? Video player ? pour tous les param?tres r?gionaux, sauf si le param?tre r?gional de l?ordinateur client est `fr-fr`, auquel cas ils verront le nom d?affichage ? Lecteur vid?o ?.

> **REMARQUE** Vous ne pouvez sp?cifier qu?un seul remplacement par langue, notamment pour les param?tres r?gionaux par d?faut. Par exemple, si votre param?tre r?gional par d?faut est `en-us`, vous ne pouvez pas sp?cifier un remplacement pour `en-us`. 

L?exemple suivant applique un remplacement de param?tre r?gional pour l??l?ment [Description]. Il commence par sp?cifier le param?tre r?gional par d?faut `en-us` et une description en anglais, puis sp?cifie une instruction [Override] avec une description en fran?ais pour le param?tre r?gional `fr-fr` :

```xml
<DefaultLocale>en-us</DefaultLocale>
...
<Description DefaultValue=
   "Watch YouTube videos referenced in the emails you receive 
   without leaving your email client.">
   <Override Locale="fr-fr" Value=
   "Visualisez les vidéos YouTube référencées dans vos courriers 
   électronique directement depuis Outlook et Outlook Web App."/>
</Description>
```

Cela signifie que le compl?ment consid?re `en-us` comme le param?tre r?gional par d?faut. Les utilisateurs verront la description en anglais figurant dans l?attribut `DefaultValue` pour tous les param?tres r?gionaux, sauf si le param?tre r?gional de l?ordinateur du client est `fr-fr`, auquel cas la description s?affichera en fran?ais.

Dans l?exemple suivant, le compl?ment sp?cifie une image s?par?e convenant mieux au param?tre r?gional et ? la culture `fr-fr`. Par d?faut, les utilisateurs voient l?image DefaultLogo.png, sauf lorsque le param?tre r?gional de l?ordinateur client est `fr-fr`. Dans ce cas, les utilisateurs voient l?image FrenchLogo.png. 


```xml
<!-- Replace "domain" with a real web server name and path. -->
<IconUrl DefaultValue="https://<domain>/DefaultLogo.png"/>
<Override Locale="fr-fr" Value="https://<domain>/FrenchLogo.png"/>
```

L?exemple suivant montre comment localiser une ressource dans la section `Resources`. Une valeur de remplacement des param?tres r?gionaux est appliqu?e pour une image plus appropri?e par rapport ? la culture `ja-jp`.

```xml
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
 ...
```


Pour l??l?ment [SourceLocation], la prise en charge de param?tres r?gionaux suppl?mentaires implique de fournir un fichier HTML source distinct pour chacun des param?tres r?gionaux sp?cifi?s. Les utilisateurs de chaque param?tre r?gional que vous sp?cifiez peuvent acc?der ? une page web personnalis?e con?ue pour eux.

Pour les compl?ments Outlook, l??l?ment [SourceLocation] s?aligne ?galement sur le facteur de forme. Cela vous permet de fournir un fichier source HTML localis? distinct pour chaque format. Vous pouvez sp?cifier un ou plusieurs ?l?ments enfant [Override] dans chaque ?l?ment de param?tres applicable ([DesktopSettings], [TabletSettings] ou [PhoneSettings]). L?exemple suivant montre les ?l?ments de param?tres pour les formats ordinateur de bureau, tablette et smartphone, avec pour chacun un fichier HTML pour le param?tre r?gional par d?faut et pour le param?tre r?gional fran?ais.


```xml
<DesktopSettings>
   <SourceLocation DefaultValue="https://contoso.com/Desktop.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Desktop.html" />
   </SourceLocation>
   <RequestedHeight>250</RequestedHeight>
</DesktopSettings>
<TabletSettings>
   <SourceLocation DefaultValue="https://contoso.com/Tablet.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Tablet.html" />
   </SourceLocation>
   <RequestedHeight>200</RequestedHeight>
</TabletSettings>
<PhoneSettings>
   <SourceLocation DefaultValue="https://contoso.com/Mobile.html">
      <Override Locale="fr-fr" Value="https://contoso.com/fr/Mobile.html" />
   </SourceLocation>
</PhoneSettings>
```

## <a name="match-datetime-format-with-client-locale"></a>Mettre en correspondance le format de date/heure avec le param?tre r?gional du client

Vous pouvez obtenir les param?tres r?gionaux de l?interface utilisateur de l?application d?h?bergement en utilisant la propri?t? [displayLanguage]. Vous pouvez ensuite afficher les valeurs de date et d?heure dans un format coh?rent avec les param?tres r?gionaux actuels de l?application h?te. Une solution consiste ? pr?parer un fichier de ressources qui sp?cifie le format d?affichage de date/heure ? utiliser pour chaque param?tre r?gional pris en charge par le compl?ment Office. Lors de l?ex?cution, votre compl?ment peut utiliser le fichier de ressources et faire correspondre le format de date/heure appropri? avec le param?tre r?gional obtenu ? partir de la propri?t? [displayLanguage].

Vous pouvez obtenir les param?tres r?gionaux des donn?es de l?application d?h?bergement en utilisant la propri?t? [contentLanguage]. En fonction de cette valeur, vous pouvez correctement interpr?ter ou afficher des cha?nes de date/heure. Par exemple, dans le param?tre r?gional `jp-JP`, les valeurs de date/heure sont exprim?es sous la forme `yyyy/MM/dd`, alors qu?avec le param?tre r?gional `fr-FR` elles apparaissent sous la forme `dd/MM/yyyy`.


## <a name="use-ajax-for-globalization-and-localization"></a>Utiliser Ajax pour l?internationalisation et la localisation


Si vous utilisez Visual Studio pour cr?er des Compl?ments Office, .NET Framework et Ajax offrent des moyens d?internationaliser et de localiser les fichiers de script client.

Vous pouvez internationaliser et utiliser les extensions de type JavaScript [Date](http://msdn.microsoft.com/library/caf98d32-2de2-4704-8198-692350343681.aspx) et [Number](http://msdn.microsoft.com/library/c216d3a1-12ae-47d1-bca1-c3666d04572f.aspx) ainsi que l?objet JavaScript [Date](http://msdn.microsoft.com/library/ce2202bb-7ec9-4f5a-bf48-3a04feff283e.aspx) dans le code JavaScript pour qu?une Compl?ment Office affiche les valeurs en fonction des param?tres r?gionaux du navigateur actuel. Pour plus d?informations, voir [Walkthrough: Globalizing a Date by Using Client Script](http://msdn.microsoft.com/library/69b34e6d-d590-4d03-a763-b7ae54b47d74.aspx).

Vous pouvez inclure des cha?nes de ressources localis?es directement dans des fichiers JavaScript autonomes pour fournir des fichiers de script client pour les diff?rents param?tres r?gionaux, qui sont d?finis dans le navigateur ou fournis par l?utilisateur. Cr?ez un fichier de script distinct pour chaque param?tre r?gional pris en charge. Dans chaque fichier de script, incluez un objet au format JSON contenant les cha?nes de ressources pour ce param?tre r?gional. Les valeurs localis?es sont appliqu?es lorsque le script s?ex?cute dans le navigateur. 


## <a name="example-build-a-localized-office-add-in"></a>Exemple : cr?er un compl?ment Office localis?

Cette section inclut des exemples expliquant comment localiser la description, le nom d?affichage et l?interface utilisateur d?une Compl?ment Office.

Pour ex?cuter l?exemple de code fourni, configurez Microsoft Office 2013 sur votre ordinateur pour utiliser des langues suppl?mentaires et pouvoir tester votre compl?ment en basculant d?une langue ? l?autre pour l?affichage des menus et des commandes, l??dition et la v?rification, ou les deux.

En outre, vous devez cr?er un projet de compl?ment Office Visual Studio 2015.

> **REMARQUE** Pour t?l?charger Visual Studio 2015, consultez la [page d?di?e aux outils de d?veloppement Office](https://www.visualstudio.com/features/office-tools-vs). Cette page contient ?galement un lien pour t?l?charger les outils de d?veloppement Office.

### <a name="configure-office-2013-to-use-additional-languages-for-display-or-editing"></a>Configurer Office 2013 pour utiliser des langues suppl?mentaires pour l?affichage ou l??dition

Vous pouvez utiliser un module linguistique Office 2013 pour installer des langues suppl?mentaires. Pour plus d?informations sur les modules linguistiques et comment les obtenir, voir [Options de langue Office 2013](http://office.microsoft.com/en-us/language-packs/).

> **REMARQUE** Si vous ?tes abonn? ? MSDN, les modules linguistiques Office 2013 peuvent ?tre disponibles dans le cadre de votre abonnement. Pour savoir si votre abonnement propose le t?l?chargement des modules linguistiques Office 2013, acc?dez ? [Accueil Abonnements MSDN](https://msdn.microsoft.com/subscriptions/manage/), tapez ? Modules linguistiques Office 2013 ? dans **T?l?chargements logiciels**, choisissez **Rechercher**, puis s?lectionnez **Produits disponibles avec mon abonnement**. Sous **Langue**, cochez la case correspondant au module linguistique que vous voulez t?l?charger, puis cliquez sur **OK**. 

Une fois le module linguistique install?, vous pouvez configurer Office 2013 pour utiliser la langue install?e pour l?affichage de l?interface utilisateur, pour l??dition du contenu du document, ou les deux. Dans cet exemple, le module linguistique espagnol a ?t? install? sur Office 2013.

### <a name="create-an-office-add-in-project"></a>Cr?er un projet de compl?ment Office

1. Dans Visual Studio, choisissez **Fichier** > **Nouveau projet**.
    
2. Dans la bo?te de dialogue **Nouveau projet**, sous **Mod?les**, d?veloppez **Visual Basic** ou **Visual C#**, d?veloppez **Office/SharePoint**, puis s?lectionnez **Compl?ments Office**.
    
3. Choisissez **Compl?ment Office** et donnez un nom ? votre compl?ment, par exemple WorldReadyApp. Cliquez sur **OK**.
    
4. Dans la bo?te de dialogue **Cr?er un compl?ment Office**, s?lectionnez **Volet Office** et cliquez sur **Suivant**. Sur la page suivante, d?sactivez les cases ? cocher pour toutes les applications h?tes ? l?exception de Word. Cliquez sur **Terminer** pour cr?er le projet.
    

### <a name="localize-the-text-used-in-your-add-in"></a>Localiser le texte utilis? dans votre compl?ment

Le texte que vous souhaitez localiser dans une autre langue appara?t ? deux emplacements :

-  **Nom d?affichage et description du compl?ment**. Ce contenu est contr?l? par les entr?es du fichier manifeste de l?application.
    
-  **Interface utilisateur du compl?ment**. Vous pouvez localiser les cha?nes qui s?affichent dans l?interface utilisateur de votre compl?ment ? l?aide du code JavaScript, par exemple en utilisant un fichier de ressources s?par? qui contient les cha?nes localis?es.
    
Pour localiser le nom d?affichage et la description du compl?ment

1. Dans l? **Explorateur de solutions**, d?veloppez **WorldReadyApp**, **WorldReadyAppManifest**, puis choisissez **WorldReadyApp.xml**.
    
2. Dans WorldReadyAppManifest.xml, remplacez les ?l?ments [DisplayName] et [Description] par le bloc de code suivant :
    
    > **REMARQUE** Vous pouvez remplacer les cha?nes localis?es en espagnol utilis?es dans cet exemple pour les ?l?ments [DisplayName] et [Description] par les cha?nes localis?es dans une autre langue.

    ```xml
    <DisplayName DefaultValue="World Ready add-in">
      <Override Locale="es-es" Value="Aplicación de uso internacional"/>
    </DisplayName>
    <Description DefaultValue="An add-in for testing localization">
      <Override Locale="es-es" Value="Una aplicación para la prueba de la localización"/>
    </Description>
    ```

3. Lorsque vous modifiez la langue d?affichage dans Office 2013, par exemple de l?anglais vers l?espagnol, puis que vous ex?cutez le compl?ment, le nom d?affichage et la description du compl?ment sont affich?s avec le texte localis?. 
    
Pour mettre en page l?interface utilisateur du compl?ment :

1. Dans Visual Studio, dans l?**Explorateur de solutions**, choisissez  **Home.html**.
    
2. Remplacez le code HTML dans Home.html par le code HTML suivant.
    
    ```html
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title></title>
        <script src="../../Scripts/jquery-1.8.2.js" type="text/javascript"></script>
    
        <link href="../../Content/Office.css" rel="stylesheet" type="text/css" />
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    
        <!-- To enable offline debugging using a local reference to Office.js, use:                        -->
        <!-- <script src="../../Scripts/Office/MicrosoftAjax.js" type="text/javascript"></script>          -->
        <!--    <script src="../../Scripts/Office/1.0/office.js" type="text/javascript"></script>          -->
    
        <link href="../App.css" rel="stylesheet" type="text/css" />
        <script src="../App.js" type="text/javascript"></script>
    
        <link href="Home.css" rel="stylesheet" type="text/css" />
        <script src="Home.js" type="text/javascript"></script> <body>
        <!-- Page content -->
        <div id="content-header">
            <div class="padding">
                <h1 id="greeting"></h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <div>
                    <p id="about"></p>
                </div>            
            </div>
        </div>
    </head>
    </html>
    ```

3. Dans Visual Studio, choisissez  **Fichier**,  **Enregistrer App\Home\Home.html**.
    
La figure suivante montre l??l?ment titre (h1) et l??l?ment paragraphe (p) qui afficheront le texte localis? lors de l?ex?cution de l?exemple de compl?ment.

*Figure 1. Interface utilisateur du compl?ment*

![Interface utilisateur de l?application avec des sections en surbrillance](../images/office15-app-how-to-localize-fig03.png)

### <a name="add-the-resource-file-that-contains-the-localized-strings"></a>Ajouter le fichier de ressources qui contient les cha?nes localis?es

Le fichier de ressources JavaScript contient les cha?nes utilis?es pour l?interface utilisateur du compl?ment. L?interface utilisateur de l?exemple de compl?ment comprend un ?l?ment h1 qui affiche un message de bienvenue et un ?l?ment p qui pr?sente le compl?ment ? l?utilisateur. 

Pour activer les cha?nes localis?es pour le titre et le paragraphe, placez les cha?nes dans un fichier de ressources distinct. Le fichier de ressources cr?e un objet JavaScript qui contient un objet JavaScript Object Notation (JSON) individuel pour chaque ensemble de cha?nes localis?es. Le fichier de ressources fournit une m?thode pour obtenir l?objet JSON appropri? pour des param?tres r?gionaux donn?s. 

Pour ajouter le fichier de ressources au projet de compl?ment :

1. Dans l?**Explorateur de solutions** de Visual Studio, s?lectionnez le dossier **Compl?ment** dans le projet web pour l?exemple de compl?ment et choisissez **Ajouter** > **Fichier JavaScript**.
    
2. Dans la bo?te de dialogue **Sp?cifier le nom de l??l?ment**, saisissez UIStrings.js.
    
3. Ajoutez le code suivant au fichier UIStrings.js.

    ```js
    /* Store the locale-specific strings */
    
    var UIStrings = (function ()
    {
        "use strict";
    
        var UIStrings = {};
    
        // JSON object for English strings
        UIStrings.EN =
        {        
            "Greeting": "Welcome",
            "Introduction": "This is my localized add-in."        
        };
    
        // JSON object for Spanish strings
        UIStrings.ES =
        {        
            "Greeting": "Bienvenido",
            "Introduction": "Esta es mi aplicación localizada."
        };
    
        UIStrings.getLocaleStrings = function (locale)
        {
            var text;
            
            // Get the resource strings that match the language.
            switch (locale)
            {
                case 'en-US':
                    text = UIStrings.EN;
                    break;
                case 'es-ES':
                    text = UIStrings.ES;
                    break;
                default:
                    text = UIStrings.EN;
                    break;
            }
    
            return text;
        };
    
        return UIStrings;
    })();
    ```

Le fichier de ressources UIStrings.js cr?e un objet **UIStrings** qui contient les cha?nes localis?es pour l?interface utilisateur de votre compl?ment. 

### <a name="localize-the-text-used-for-the-add-in-ui"></a>Localiser le texte utilis? pour l?interface utilisateur du compl?ment

Pour utiliser le fichier de ressources de votre compl?ment, vous devez ajouter une balise de script pour ce fichier dans Home.html. Quand Home.html est charg?, UIStrings.js s?ex?cute et l?objet  **UIStrings** que vous utilisez pour obtenir les cha?nes est disponible pour votre code. Ajoutez le code HTML suivant dans la balise head pour Home.html pour que **UIStrings** soit disponible pour votre code.

```html
<!-- Resource file for localized strings:                                                          -->
<script src="../UIStrings.js" type="text/javascript"></script>
```

Vous pouvez d?sormais utiliser l?objet **UIStrings** pour d?finir les cha?nes pour l?interface utilisateur de votre compl?ment.

Si vous voulez changer la localisation pour votre compl?ment en fonction de la langue utilis?e pour afficher les menus et les commandes dans l?application h?te, utilisez la propri?t? **Office.context.displayLanguage** pour obtenir les param?tres r?gionaux pour cette langue. Par exemple, si la langue de l?application h?te utilise l?espagnol pour afficher les menus et les commandes, la propri?t? **Office.context.displayLanguage** retournera le code de langue es-ES.

Si vous voulez changer la localisation pour votre compl?ment en fonction de la langue utilis?e pour l??dition du contenu de document, utilisez la propri?t?  **Office.context.contentLanguage** pour obtenir les param?tres r?gionaux pour cette langue. Par exemple, si la langue de l?application h?te utilise l?espagnol pour l??dition de contenu de document, la propri?t? **Office.context.contentLanguage** retournera le code de langue es-ES.

Une fois que vous connaissez la langue utilis?e par l?application h?te, vous pouvez utiliser **UIStrings** pour obtenir les cha?nes localis?es qui correspondent ? la langue de l?application h?te.

Remplacez le code du fichier Home.js par le code suivant. Le code montre comment changer les cha?nes utilis?es dans les ?l?ments d?interface utilisateur de Home.html en fonction de la langue d?affichage de l?application h?te ou de la langue d??dition de l?application h?te.

> **REMARQUE** Pour activer ou d?sactiver la localisation du compl?ment en fonction de la langue utilis?e pour la modification, supprimez le commentaire de la ligne de code `var myLanguage = Office.context.contentLanguage;` et ajoutez un commentaire ? la ligne de code `var myLanguage = Office.context.displayLanguage;`

```js
/// <reference path="../App.js" />
/// <reference path="../UIStrings.js" />


(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason)
    {
       
        $(document).ready(function () {
            app.initialize();

            // Get the language setting for editing document content.
            // To test this, uncomment the following line and then comment out the
            // line that uses Office.context.displayLanguage.
            // var myLanguage = Office.context.contentLanguage;

            // Get the language setting for UI display in the host application.
            var myLanguage = Office.context.displayLanguage;            
            var UIText;

            // Get the resource strings that match the language.
            // Use the UIStrings object from the UIStrings.js file
            // to get the JSON object with the correct localized strings.
            UIText = UIStrings.getLocaleStrings(myLanguage);            

            // Set localized text for UI elements.
            $("#greeting").text(UIText.Greeting);
            $("#about").text(UIText.Instruction);
        });
    };    
})();
```

### <a name="test-your-localized-add-in"></a>Tester votre compl?ment localis?

Pour tester votre compl?ment localis?, changez la langue utilis?e pour l?affichage et l??dition dans l?application h?te, puis ex?cutez votre compl?ment. 

Pour changer la langue utilis?e pour l?affichage ou l??dition dans votre compl?ment :

1. Dans Word 2013, s?lectionnez **Fichier** > **Options** > **Langue**. La figure suivante montre la bo?te de dialogue **Options Word** ouverte sur l?onglet Langue.
    
    *Figure 2. Options de langue dans la bo?te de dialogue Options Word 2013*

    ![Bo?te de dialogue Options Word 2013](../images/office15-app-how-to-localize-fig04.png)

2. Sous **Choisir les langues de l?interface utilisateur et de l?Aide**, s?lectionnez la langue souhait?e pour l?affichage, par exemple l?espagnol, puis cliquez sur la fl?che vers le haut pour d?placer l?espagnol tout en haut de la liste. Pour changer la langue utilis?e pour l??dition, sous **Choisir les langues d??dition**, choisissez la langue ? utiliser pour l??dition, par exemple l?espagnol, puis choisissez **D?finir par d?faut**.
    
3. S?lectionnez **OK** pour confirmer votre choix, puis fermez Word.
    
Ex?cutez l?exemple de compl?ment. Le compl?ment de volet de t?ches est charg? dans Word 2013 et les cha?nes de l?interface utilisateur du compl?ment changent pour correspondre ? la langue utilis?e par l?application h?te, comme indiqu? dans la figure suivante.


*Figure 3. Interface utilisateur du compl?ment avec le texte localis?*

![Application avec le texte de l?interface utilisateur localis?](../images/office15-app-how-to-localize-fig05.png)

## <a name="see-also"></a>Voir aussi

- [Instructions de conception pour les compl?ments Office](../design/add-in-design.md)    
- [Identificateurs de langue et valeurs d?ID de l??l?ment OptionState dans Office 2013](http://technet.microsoft.com/en-us/library/cc179219%28Office.15%29.aspx)

[DefaultLocale]:        https://dev.office.com/reference/add-ins/manifest/defaultlocale
[Description]:          https://dev.office.com/reference/add-ins/manifest/description
[DisplayName]:          https://dev.office.com/reference/add-ins/manifest/displayname
[IconUrl]:              https://dev.office.com/reference/add-ins/manifest/iconurl
[HighResolutionIconUrl]:https://dev.office.com/reference/add-ins/manifest/highresolutioniconurl
[Ressources]:            https://dev.office.com/reference/add-ins/manifest/resources
[SourceLocation]:       https://dev.office.com/reference/add-ins/manifest/sourcelocation
[Override]:             https://dev.office.com/reference/add-ins/manifest/override
[DesktopSettings]:      https://dev.office.com/reference/add-ins/manifest/desktopsettings
[TabletSettings]:       https://dev.office.com/reference/add-ins/manifest/tabletsettings
[PhoneSettings]:        https://dev.office.com/reference/add-ins/manifest/phonesettings
[displayLanguage]:  https://dev.office.com/reference/add-ins/shared/office.context.displaylanguage 
[contentLanguage]:  https://dev.office.com/reference/add-ins/shared/office.context.contentlanguage 
[RFC 3066]: https://www.rfc-editor.org/info/rfc3066
