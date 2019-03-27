---
title: Localisation des compléments Office
description: Vous pouvez utiliser l’API JavaScript pour Office afin de déterminer des paramètres régionaux et d’afficher des chaînes basées sur les paramètres régionaux de l’application hôte, ou afin d’interpréter ou d’afficher des données en fonction des paramètres régionaux des données.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: b01535b76485847783104ce840256949ce17ca26
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871709"
---
# <a name="localization-for-office-add-ins"></a>Localisation des compléments Office

Vous pouvez librement implémenter n’importe quel schéma de localisation convenant à votre Complément Office. L’API JavaScript et le schéma du manifeste de la plateforme Compléments Office offrent quelques choix. Vous pouvez utiliser l’API JavaScript pour Office pour déterminer un paramètre régional et les chaînes d’affichage en fonction des paramètres régionaux de l’application hôte, ou pour interpréter ou afficher les données en fonction des paramètres régionaux des données. Vous pouvez utiliser le manifeste pour spécifier l’emplacement des fichiers et les informations descriptives propres à un paramètre régional. Sinon, vous pouvez utiliser un script Microsoft Ajax pour prendre en charge l’internationalisation et la localisation.

## <a name="use-the-javascript-api-to-determine-locale-specific-strings"></a>Utiliser l’API JavaScript pour déterminer les chaînes propres aux paramètres régionaux

L’API JavaScript pour Office offre deux propriétés qui prennent en charge l’affichage ou l’interprétation de valeurs cohérentes avec les paramètres régionaux de l’application hôte et des données :

- [Context.displayLanguage][displayLanguage] spécifie les paramètres régionaux (ou la langue) de l’interface utilisateur de l’application hôte. L’exemple suivant vérifie si l’application hôte utilise les paramètres régionaux en-US ou fr-FR, et affiche un message de bienvenue propre aux paramètres régionaux.

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

- [Context.contentLanguage][contentLanguage] spécifie le paramètre régional (ou langue) des données. Le fait d’étendre le dernier exemple de code, au lieu de vérifier la propriété [displayLanguage], attribue la valeur`myLanguage` de la propriété [contentLanguage] et utilise le reste du code pour afficher un message de bienvenue correspondant aux paramètres régionaux des données :

    ```js
    var myLanguage = Office.context.contentLanguage;
    ```

## <a name="control-localization-from-the-manifest"></a>Contrôler la localisation à partir du manifeste


Chaque complément Office indique un élément [DefaultLocale] élément et un paramètre régional dans son manifeste. Par défaut, la plateforme de complément Office et les applications hôtes Office appliquent les valeurs des éléments [Description], [DisplayName], [IconUrl], [HighResolutionIconUrl] et [SourceLocation] à tous les paramètres régionaux. Vous pouvez éventuellement prendre en charge des valeurs spécifiques pour les paramètres régionaux spécifiques, en spécifiant un élément enfant [Override] pour chaque paramètre régional supplémentaire, pour chacun des cinq éléments. La valeur de l’élément [DefaultLocale] et de l’attribut `Locale` de l’élément [Override] est spécifiée en fonction de la norme [RFC 3066] relative aux balises pour l’identification des langues (« Tags for the Identification of Languages »). Le tableau 1 décrit la prise en charge de localisation de ces éléments.

*Tableau 1. Prise en charge de localisation*


|**Élément**|**Prise en charge de localisation**|
|:-----|:-----|
|[Description]   |Les utilisateurs de chaque paramètre régional que vous spécifiez peuvent voir une description localisée du complément dans AppSource (ou dans un catalogue privé).<br/>Pour les compléments Outlook, les utilisateurs peuvent voir la description dans le Centre d’administration Exchange (EAC) après l’installation.|
|[DisplayName]   |Les utilisateurs de chaque paramètre régional que vous spécifiez peuvent voir une description localisée du complément dans AppSource (ou dans un catalogue privé).<br/>Pour les compléments Outlook, les utilisateurs peuvent voir le nom d’affichage sous forme d’étiquette pour le bouton de l’application Outlook ainsi que dans l’EAC après l’installation.<br/>Pour les compléments de contenu et du volet Office, les utilisateurs peuvent voir l’icône dans le ruban après avoir installé l’application.|
|[IconUrl]        |L’image de l’icône est facultative. Vous pouvez utiliser la même technique de remplacement pour spécifier une image donnée pour une culture particulière. Si vous utilisez et localisez une icône, les utilisateurs de chaque paramètre régional que vous spécifiez peuvent voir l’image d’icône localisée pour le complément.<br/>Si vous utilisez et localisez une icône, les utilisateurs de chaque paramètre régional que vous spécifiez peuvent voir l’image d’icône localisée pour le complément.<br/>Pour les compléments Outlook, les utilisateurs peuvent voir l’icône dans l’EAC après l’installation du complément.|
|[HighResolutionIconUrl] **Important :** cet élément est disponible uniquement lors de l’utilisation de la version 1.1 du manifeste de complément.|L’image de l’icône de haute résolution est facultative. Néanmoins, si elle est indiquée, elle doit l’être après l’élément [IconUrl]. Si  [HighResolutionIconUrl] est spécifié et que le complément est installé sur un appareil qui prend en charge la haute résolution (dpi), la valeur [HighResolutionIconUrl] est utilisée à la place de la valeur [IconUrl].<br/>Si  HighResolutionIconUrl est spécifié et que le complément est installé sur un appareil qui prend en charge la haute résolution (dpi), la valeur HighResolutionIconUrl est utilisée à la place de la valeur IconUrl.<br/>Si vous utilisez et localisez une icône, les utilisateurs de chaque paramètre régional que vous spécifiez peuvent voir l’image d’icône localisée pour le complément.<br/>Pour les compléments Outlook, les utilisateurs peuvent voir l’icône dans l’EAC après l’installation du complément.|
|Pour les compléments de contenu et du volet de tâches, les utilisateurs peuvent voir l’icône dans le ruban après avoir installé le complément.   |[Ressources] Important : cet élément est disponible uniquement lors de l’utilisation de la version 1.1 du manifeste de complément. |
|[SourceLocation]   |Les utilisateurs de chaque paramètre régional que vous spécifiez peuvent voir une page web que vous concevez spécifiquement pour le complément pour ce paramètre régional. |


> [!NOTE]
> Vous pouvez localiser la description et le nom d’affichage uniquement pour les paramètres régionaux qu’Office prend en charge. Pour obtenir la liste des langues et les paramètres régionaux pour la version actuelle d’Office, voir [Identificateurs de langue et valeurs d’ID de l’élément OptionState dans Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15)).


### <a name="examples"></a>Exemples

Par exemple, un complément Office peut spécifier [DefaultLocale] en tant que `en-us`. Pour l’élément [DisplayName], le complément peut spécifier un élément enfant [Override] pour le paramètre régional `fr-fr`, comme illustré ci-dessous.


```xml
<DefaultLocale>en-us</DefaultLocale>
...
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

> [!NOTE]
> Si vous devez localiser plusieurs domaines au sein d’une famille de langues, comme `de-de` et `de-at`, nous vous recommandons d’utiliser des éléments `Override` distincts pour chaque domaine. L’utilisation uniquement du nom de la langue, soit `de` dans ce cas, n’est pas prise en charge pour toutes les combinaisons de plateformes et d’applications hôtes Office.

Cela signifie que le complément adopte le paramètre régional `en-us` par défaut. Les utilisateurs voient le nom d’affichage « Video player » pour tous les paramètres régionaux, sauf si le paramètre régional de l’ordinateur client est `fr-fr`, auquel cas ils verront le nom d’affichage « Lecteur vidéo ».

> [!NOTE]
> Vous ne pouvez spécifier qu’un seul remplacement par langue, notamment pour les paramètres régionaux par défaut. Par exemple, si votre paramètre régional par défaut est `en-us`, vous ne pouvez pas spécifier un remplacement pour `en-us`. 

L’exemple suivant applique un remplacement de paramètre régional pour l’élément [Description]. Il commence par spécifier le paramètre régional par défaut `en-us` et une description en anglais, puis spécifie une instruction [Override] avec une description en français pour le paramètre régional `fr-fr` :

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

Il commence par spécifier le paramètre régional par défaut `en-us` et une description en anglais, puis spécifie une instruction `DefaultValue` avec une description en français pour le paramètre régional `fr-fr`:

Les utilisateurs verront la description en anglais figurant dans l’attribut `fr-fr` pour tous les paramètres régionaux, sauf si le paramètre régional de l’ordinateur du client est `fr-fr`, auquel cas la description s’affichera en français. 


```xml
<!-- Replace "domain" with a real web server name and path. -->
<IconUrl DefaultValue="https://<domain>/DefaultLogo.png"/>
<Override Locale="fr-fr" Value="https://<domain>/FrenchLogo.png"/>
```

Dans ce cas, les utilisateurs voient l’image FrenchLogo.png.

```xml
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
 ...
```


Une valeur de remplacement des paramètres régionaux est appliquée pour une image plus appropriée par rapport à la culture [].

Pour les compléments Outlook, l’élément [SourceLocation] s’aligne également sur le facteur de forme. Cela vous permet de fournir un fichier source HTML localisé distinct pour chaque format. Vous pouvez spécifier un ou plusieurs éléments enfant [Override] dans chaque élément de paramètres applicable ([DesktopSettings], [TabletSettings] ou [PhoneSettings]). L’exemple suivant montre les éléments de paramètres pour les formats ordinateur de bureau, tablette et smartphone, avec pour chacun un fichier HTML pour le paramètre régional par défaut et pour le paramètre régional français.


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

## <a name="match-datetime-format-with-client-locale"></a>Mettre en correspondance le format de date/heure avec le paramètre régional du client

Vous pouvez obtenir les paramètres régionaux de l’interface utilisateur de l’application d’hébergement en utilisant la propriété [displayLanguage]. Vous pouvez ensuite afficher les valeurs de date et d’heure dans un format cohérent avec les paramètres régionaux actuels de l’application hôte. Une solution consiste à préparer un fichier de ressources qui spécifie le format d’affichage de date/heure à utiliser pour chaque paramètre régional pris en charge par le complément Office. Lors de l’exécution, votre complément peut utiliser le fichier de ressources et faire correspondre le format de date/heure approprié avec le paramètre régional obtenu à partir de la propriété [displayLanguage].

Vous pouvez obtenir les paramètres régionaux des données de l’application d’hébergement en utilisant la propriété [contentLanguage]. En fonction de cette valeur, vous pouvez correctement interpréter ou afficher des chaînes de date/heure. Par exemple, dans le paramètre régional `jp-JP`, les valeurs de date/heure sont exprimées sous la forme `yyyy/MM/dd`, alors qu’avec le paramètre régional `fr-FR` elles apparaissent sous la forme `dd/MM/yyyy`.


## <a name="use-ajax-for-globalization-and-localization"></a>Utiliser Ajax pour l’internationalisation et la localisation


Si vous utilisez Visual Studio pour créer des Compléments Office, .NET Framework et Ajax offrent des moyens d’internationaliser et de localiser les fichiers de script client.

Si vous utilisez Visual Studio pour créer des Compléments Office, .NET Framework et Ajax offrent des moyens d’internationaliser et de localiser les fichiers de script client.

Pour plus d’informations, voir Walkthrough: Globalizing a Date by Using Client Script.


## <a name="example-build-a-localized-office-add-in"></a>Exemple : créer un complément Office localisé

Cette section inclut des exemples expliquant comment localiser la description, le nom d’affichage et l’interface utilisateur d’une Complément Office. 

> [!NOTE]
> Pour télécharger Visual Studio 2017, voir la [Page Visual Studio IDE](https://visualstudio.microsoft.com/vs/). Lors de l’installation, vous devez sélectionner la charge de travail de développement Office/SharePoint.

### <a name="configure-office-to-use-additional-languages-for-display-or-editing"></a>Configurer Office pour utiliser des langues supplémentaires pour l’affichage ou l’édition

Pour exécuter l’exemple de code fourni, configurez Microsoft Office sur votre ordinateur pour utiliser des langues supplémentaires pour pouvoir tester votre complément en basculant d’une langue à l’autre pour l’affichage des menus et des commandes, l’édition et la vérification, ou les deux.

Vous pouvez utiliser un module linguistique Office pour installer une autre langue. Pour plus d’informations sur les Modules linguistiques et où les obtenir, voir [Pack d’accessoires linguistiques pour Office](https://office.microsoft.com/language-packs/).

Après avoir installé le Pack d’accessoires linguistiques, vous pouvez configurer Office pour utiliser la langue installée pour l’affichage dans l’interface utilisateur, pour modifier du contenu de document, ou les deux. L’exemple dans cet article utilise une installation d’Office qui contient le module linguistique espagnol.

### <a name="create-an-office-add-in-project"></a>Créer un projet de complément Office

Vous devez créer un projet de complément Office Visual Studio 2017.

> [!NOTE]
> Si vous n’avez pas installé Visual Studio 2017, voir la [Page Visual Studio IDE](https://visualstudio.microsoft.com/vs/) pour obtenir des instructions de téléchargement. Lors de l’installation, vous devez sélectionner la charge de travail de développement Office/SharePoint. Si vous avez déjà installé Visual Studio 2017, [utilisez Visual Studio Installer](/visualstudio/install/modify-visual-studio/) pour vérifier que la charge de travail de développement Office/SharePoint est bien installée.


1. Dans Visual Studio, choisissez **Fichier** > **Nouveau projet**.
2. Dans la boîte de dialogue **Nouveau projet**, développez **Visual Basic** ou **Visual C#**, développez **Office/SharePoint**, puis sélectionnez **Compléments**.
3. Choisissez **Complément Word** et nommez votre complément **WorldReadyAddIn**. Cliquez sur **OK**.

### <a name="localize-the-text-used-in-your-add-in"></a>Localiser le texte utilisé dans votre complément

Le texte que vous souhaitez localiser dans une autre langue apparaît à deux emplacements :

-  **Nom d’affichage et description du complément**. Ce contenu est contrôlé par les entrées du fichier manifeste de l’application.

-  **Interface utilisateur du complément**. Vous pouvez localiser les chaînes qui s’affichent dans l’interface utilisateur de votre complément à l’aide du code JavaScript, par exemple en utilisant un fichier de ressources séparé contenant les chaînes localisées.

Pour localiser le nom d’affichage et la description du complément :

1. Pour localiser le nom d’affichage et la description du complément

2. Dans WorldReadyAppManifest.xml, remplacez les éléments [DisplayName] et [Description] par le bloc de code suivant :

    > [!NOTE]
    > Vous pouvez remplacer les chaînes localisées en espagnol utilisées dans cet exemple pour les éléments [DisplayName] et [Description] par les chaînes localisées en une autre langue.

    ```xml
    <DisplayName DefaultValue="World Ready add-in">
      <Override Locale="es-es" Value="Aplicación de uso internacional"/>
    </DisplayName>
    <Description DefaultValue="An add-in for testing localization">
      <Override Locale="es-es" Value="Una aplicación para la prueba de la localización"/>
    </Description>
    ```

3. Lorsque vous modifiez la langue d’affichage dans Office 2013, par exemple de l’anglais vers l’espagnol, puis que vous exécutez le complément, le nom d’affichage et la description du complément sont affichés avec le texte localisé.

Lorsque vous modifiez la langue d’affichage dans Office 2013, par exemple de l’anglais vers l’espagnol, puis que vous exécutez le complément, le nom d’affichage et la description du complément sont affichés avec le texte localisé.

1. Dans Visual Studio, dans l’**Explorateur de solutions**, choisissez **Home.html**.

2. Remplacez le contenu de l’élément `<body>` dans Home.html par le HTML suivant et enregistrez le fichier.

    ```html
    <body>
        <!-- Page content -->
        <div id="content-header" class="ms-bgColor-themePrimary ms-font-xl">
            <div class="padding">
                <h1 id="greeting" class="ms-fontColor-white"></h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <div class="ms-font-m">
                    <p id="about"></p>
                </div>
            </div>
        </div>
    </body>
    ```

La figure suivante montre l’élément titre (h1) et l’élément paragraphe (p) qui afficheront le texte localisé lorsque vous terminez les étapes restantes et exécutez le complément.

*Figure 1. Interface utilisateur du complément*

![Interface utilisateur de l’application avec des sections en surbrillance](../images/office15-app-how-to-localize-fig03.png)

### <a name="add-the-resource-file-that-contains-the-localized-strings"></a>Ajouter le fichier de ressources qui contient les chaînes localisées

Le fichier de ressource JavaScript contient les chaînes utilisées pour l’interface utilisateur du complément. Le code HTML pour l’exemple d’interface utilisateur du complément contient un `<h1>` élément qui affiche un message d’accueil et un `<p>` élément qui présente le complément à l’utilisateur. 

Pour activer les chaînes localisées pour le titre et le paragraphe, placez les chaînes dans un fichier de ressources distinct. Le fichier de ressources crée un objet JavaScript qui contient un objet JavaScript Object Notation (JSON) individuel pour chaque ensemble de chaînes localisées. Le fichier de ressources fournit une méthode pour obtenir l’objet JSON approprié pour des paramètres régionaux donnés.

Pour ajouter le fichier de ressources au projet de complément :

1. Dans **l’Explorateur de solutions** dans Visual Studio, cliquez avec le bouton droit sur le projet**WorldReadyAddInWeb**, puis choisissez **Ajouter** > **Nouvel élément**. 

2. Dans la boîte de dialogue**Ajouter un nouvel élément**, choisissez **Fichier JavaScript**.

3. Entrez **UIStrings.js** comme nom de fichier puis sélectionnez **Ajouter**.

4. Ajoutez le code suivant au fichier UIStrings.js et enregistrez le fichier.

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

Le fichier de ressources UIStrings.js crée un objet **UIStrings** qui contient les chaînes localisées pour l’interface utilisateur de votre complément.

### <a name="localize-the-text-used-for-the-add-in-ui"></a>Localiser le texte utilisé pour l’interface utilisateur du complément

Pour utiliser le fichier de ressources de votre complément, vous devez ajouter une balise de script pour ce fichier dans Home.html. Quand Home.html est chargé, UIStrings.js s’exécute et l’objet  **UIStrings** que vous utilisez pour obtenir les chaînes est disponible pour votre code. Ajoutez le code HTML suivant dans la balise head pour Home.html pour que **UIStrings** soit disponible pour votre code.

```html
<!-- Resource file for localized strings:                                                          -->
<script src="../UIStrings.js" type="text/javascript"></script>
```

Ajoutez le code HTML suivant dans la balise head pour Home.html pour que **UIStrings** soit disponible pour votre code.

Vous pouvez désormais utiliser l’objet **UIStrings** pour définir les chaînes pour l’interface utilisateur de votre complément.

Par exemple, si la langue de l’application hôte utilise l’espagnol pour afficher les menus et les commandes, la propriété **Office.context.displayLanguage** retournera le code de langue es-ES.

Par exemple, si la langue de l’application hôte utilise l’espagnol pour l’édition de contenu de document, la propriété **Office.context.contentLanguage** retournera le code de langue es-ES.

Remplacez le code du fichier Home.js par le code suivant. Le code montre comment changer les chaînes utilisées dans les éléments d’interface utilisateur de Home.html en fonction de la langue d’affichage de l’application hôte ou de la langue d’édition de l’application hôte.

> [!NOTE]
> Pour activer ou désactiver la localisation du complément en fonction de la langue utilisée pour l’édition, supprimez le commentaire de la ligne de code `var myLanguage = Office.context.contentLanguage;` et ajoutez un commentaire à la ligne de code `var myLanguage = Office.context.displayLanguage;`

```js
/// <reference path="../App.js" />
/// <reference path="../UIStrings.js" />


(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason)
    {

        $(document).ready(function () {
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
            $("#about").text(UIText.Introduction);
        });
    };
})();
```

### <a name="test-your-localized-add-in"></a>Tester votre complément localisé

Pour tester votre complément localisé, changez la langue utilisée pour l’affichage et l’édition dans l’application hôte, puis exécutez votre complément.

Pour changer la langue utilisée pour l’affichage ou l’édition dans votre complément :

1. Dans Word, sélectionnez **Fichier** > **Options** > **Langue**. La figure suivante montre la boîte de dialogue **Options Word** ouverte sous l’onglet Langue.

    *Figure 2. Options de langue dans la boîte de dialogue Options Word*

    ![Boîte de dialogue Options Word](../images/office15-app-how-to-localize-fig04.png)

2. Sous **Choisir la langue d’affichage**, sélectionnez la langue que vous souhaitez afficher, par exemple espagnol, puis sélectionnez la flèche vers le haut pour déplacer la langue Espagnol en première position dans la liste. Vous pouvez également modifier la langue utilisée pour l’édition, sous **Choisir les langues d’édition**, sélectionnez la langue que vous voulez utiliser pour l’édition, par exemple, espagnol, puis **Définir par défaut**.

3. Sélectionnez **OK** pour confirmer votre choix, puis fermez Word.

4. Appuyez sur **F5** dans Visual Studio pour exécuter le complément d’exemple ou choisissez **Déboguer** > **Démarrer le débogage** dans la barre de menus.

5. Dans Word, sélectionnez **Accueil** > **Afficher le volet de tâches**.

Lors de l’exécution, les chaînes de l’interface utilisateur du complément changent pour correspondre à la langue utilisée par l’application hôte, comme indiqué dans la figure suivante.


*Figure 3. Interface utilisateur du complément avec texte localisé*

![Application avec le texte de l’interface utilisateur localisé](../images/office15-app-how-to-localize-fig05.png)

## <a name="see-also"></a>Voir aussi

- [Instructions de conception pour les compléments Office](../design/add-in-design.md)
- [Instructions de conception pour les compléments Office](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))

[DefaultLocale]:        https://docs.microsoft.com/office/dev/add-ins/reference/manifest/defaultlocale
[Description]:          https://docs.microsoft.com/office/dev/add-ins/reference/manifest/description
[DisplayName]:          https://docs.microsoft.com/office/dev/add-ins/reference/manifest/displayname
[IconUrl]:              https://docs.microsoft.com/office/dev/add-ins/reference/manifest/iconurl
[HighResolutionIconUrl]:https://docs.microsoft.com/office/dev/add-ins/reference/manifest/highresolutioniconurl
[Resources]:            https://docs.microsoft.com/office/dev/add-ins/reference/manifest/resources
[SourceLocation]:       https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation
[Override]:             https://docs.microsoft.com/office/dev/add-ins/reference/manifest/override
[DesktopSettings]:      https://docs.microsoft.com/office/dev/add-ins/reference/manifest/desktopsettings
[TabletSettings]:       https://docs.microsoft.com/office/dev/add-ins/reference/manifest/tabletsettings
[PhoneSettings]:        https://docs.microsoft.com/office/dev/add-ins/reference/manifest/phonesettings
[displayLanguage]:  https://docs.microsoft.com/javascript/api/office/office.context#displaylanguage 
[contentLanguage]:  https://docs.microsoft.com/javascript/api/office/office.context#contentlanguage 
[RFC 3066]: https://www.rfc-editor.org/info/rfc3066
