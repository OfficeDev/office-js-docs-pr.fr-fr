---
title: Localisation des compléments Office
description: Utilisez l’API JavaScript Office pour déterminer un paramètre local et afficher des chaînes en fonction des paramètres régionaux de l’application Office, ou pour interpréter ou afficher des données en fonction des paramètres régionaux des données.
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: b49d64f2c9391539ac2d5929ebff2a4ecc08b630
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349825"
---
# <a name="localization-for-office-add-ins"></a><span data-ttu-id="f0124-103">Localisation des compléments Office</span><span class="sxs-lookup"><span data-stu-id="f0124-103">Localization for Office Add-ins</span></span>

<span data-ttu-id="f0124-104">Vous pouvez librement implémenter n’importe quel schéma de localisation convenant à votre Complément Office.</span><span class="sxs-lookup"><span data-stu-id="f0124-104">You can implement any localization scheme that's appropriate for your Office Add-in.</span></span> <span data-ttu-id="f0124-105">L’API JavaScript et le schéma du manifeste de la plateforme Compléments Office offrent quelques choix.</span><span class="sxs-lookup"><span data-stu-id="f0124-105">The JavaScript API and manifest schema of the Office Add-ins platform provide some choices.</span></span> <span data-ttu-id="f0124-106">Vous pouvez utiliser l’API JavaScript Office pour déterminer un paramètre local et afficher des chaînes en fonction des paramètres régionaux de l’application Office, ou pour interpréter ou afficher des données en fonction des paramètres régionaux des données.</span><span class="sxs-lookup"><span data-stu-id="f0124-106">You can use the Office JavaScript API to determine a locale and display strings based on the locale of the Office application, or to interpret or display data based on the locale of the data.</span></span> <span data-ttu-id="f0124-107">Vous pouvez utiliser le manifeste pour spécifier l’emplacement des fichiers et les informations descriptives propres à un paramètre régional.</span><span class="sxs-lookup"><span data-stu-id="f0124-107">You can use the manifest to specify locale-specific add-in file location and descriptive information.</span></span> <span data-ttu-id="f0124-108">Sinon, vous pouvez utiliser un script Microsoft Ajax pour prendre en charge l’internationalisation et la localisation.</span><span class="sxs-lookup"><span data-stu-id="f0124-108">Alternatively, you can use Microsoft Ajax script to support globalization and localization.</span></span>

## <a name="use-the-javascript-api-to-determine-locale-specific-strings"></a><span data-ttu-id="f0124-109">Utiliser l’API JavaScript pour déterminer les chaînes propres aux paramètres régionaux</span><span class="sxs-lookup"><span data-stu-id="f0124-109">Use the JavaScript API to determine locale-specific strings</span></span>

<span data-ttu-id="f0124-110">L Office API JavaScript fournit deux propriétés qui assurent l’affichage ou l’interprétation de valeurs cohérentes avec les paramètres régionaux de l’application Office données :</span><span class="sxs-lookup"><span data-stu-id="f0124-110">The Office JavaScript API provides two properties that support displaying or interpreting values consistent with the locale of the Office application and data:</span></span>

- <span data-ttu-id="f0124-111">[Context.displayLanguage][displayLanguage] spécifie les paramètres régionaux (ou la langue) de l’interface utilisateur de l Office application.</span><span class="sxs-lookup"><span data-stu-id="f0124-111">[Context.displayLanguage][displayLanguage] specifies the locale (or language) of the user interface of the Office application.</span></span> <span data-ttu-id="f0124-112">L’exemple suivant vérifie si l’application Office utilise les paramètres régionaux en-US ou fr-FR et affiche un message d’accueil spécifique aux paramètres régionaux.</span><span class="sxs-lookup"><span data-stu-id="f0124-112">The following example verifies if the Office application uses the en-US or fr-FR locale, and displays a locale-specific greeting.</span></span>

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

- <span data-ttu-id="f0124-p103">[Context.contentLanguage][contentLanguage] spécifie le paramètre régional (ou langue) des données. Le fait d’étendre le dernier exemple de code, au lieu de vérifier la propriété [displayLanguage], attribue la valeur`myLanguage` de la propriété [contentLanguage] et utilise le reste du code pour afficher un message de bienvenue correspondant aux paramètres régionaux des données :</span><span class="sxs-lookup"><span data-stu-id="f0124-p103">[Context.contentLanguage][contentLanguage] specifies the locale (or language) of the data. Extending the last code sample, instead of checking the [displayLanguage] property, assign `myLanguage` the value of the [contentLanguage] property, and use the rest of the same code to display a greeting based on the locale of the data:</span></span>

    ```js
    var myLanguage = Office.context.contentLanguage;
    ```

## <a name="control-localization-from-the-manifest"></a><span data-ttu-id="f0124-115">Contrôler la localisation à partir du manifeste</span><span class="sxs-lookup"><span data-stu-id="f0124-115">Control localization from the manifest</span></span>


<span data-ttu-id="f0124-116">Chaque complément Office indique un élément [DefaultLocale] élément et un paramètre régional dans son manifeste.</span><span class="sxs-lookup"><span data-stu-id="f0124-116">Every Office Add-in specifies a [DefaultLocale] element and a locale in its manifest.</span></span> <span data-ttu-id="f0124-117">Par défaut, la plateforme du Office et les applications clientes Office appliquent les valeurs des éléments [Description,] [DisplayName,] [IconUrl,] [HighResolutionIconUrl]et [SourceLocation] à tous les paramètres régionaux.</span><span class="sxs-lookup"><span data-stu-id="f0124-117">By default, the Office Add-in platform and Office client applications apply the values of the [Description], [DisplayName], [IconUrl], [HighResolutionIconUrl], and [SourceLocation] elements to all locales.</span></span> <span data-ttu-id="f0124-118">Vous pouvez éventuellement prendre en charge des valeurs spécifiques pour les paramètres régionaux spécifiques, en spécifiant un élément enfant [Override] pour chaque paramètre régional supplémentaire, pour chacun des cinq éléments.</span><span class="sxs-lookup"><span data-stu-id="f0124-118">You can optionally support specific values for specific locales, by specifying an [Override] child element for each additional locale, for any of these five elements.</span></span> <span data-ttu-id="f0124-119">La valeur de l’élément [DefaultLocale] et de l’attribut `Locale` de l’élément [Override] est spécifiée en fonction de la norme [RFC 3066] relative aux balises pour l’identification des langues (« Tags for the Identification of Languages »).</span><span class="sxs-lookup"><span data-stu-id="f0124-119">The value for the [DefaultLocale] element and for the `Locale` attribute of the [Override] element is specified according to [RFC 3066], "Tags for the Identification of Languages."</span></span> <span data-ttu-id="f0124-120">Le tableau 1 décrit la prise en charge de localisation de ces éléments.</span><span class="sxs-lookup"><span data-stu-id="f0124-120">Table 1 describes the localizing support for these elements.</span></span>

<span data-ttu-id="f0124-121">*Tableau 1. Prise en charge de localisation*</span><span class="sxs-lookup"><span data-stu-id="f0124-121">*Table 1. Localization support*</span></span>


|<span data-ttu-id="f0124-122">**Élément**</span><span class="sxs-lookup"><span data-stu-id="f0124-122">**Element**</span></span>|<span data-ttu-id="f0124-123">**Prise en charge de localisation**</span><span class="sxs-lookup"><span data-stu-id="f0124-123">**Localization support**</span></span>|
|:-----|:-----|
|<span data-ttu-id="f0124-124">[Description]</span><span class="sxs-lookup"><span data-stu-id="f0124-124">[Description]</span></span>   |<span data-ttu-id="f0124-125">Les utilisateurs de chaque paramètre régional que vous spécifiez peuvent voir une description localisée du complément dans AppSource (ou dans un catalogue privé).</span><span class="sxs-lookup"><span data-stu-id="f0124-125">Users in each locale you specify can see a localized description for the add-in in AppSource (or private catalog).</span></span><br/><span data-ttu-id="f0124-126">Pour les compléments Outlook, les utilisateurs peuvent voir la description dans le Centre d’administration Exchange (EAC) après l’installation.</span><span class="sxs-lookup"><span data-stu-id="f0124-126">For Outlook add-ins, users can see the description in the Exchange Admin Center (EAC) after installation.</span></span>|
|<span data-ttu-id="f0124-127">[DisplayName]</span><span class="sxs-lookup"><span data-stu-id="f0124-127">[DisplayName]</span></span>   |<span data-ttu-id="f0124-128">Les utilisateurs de chaque paramètre régional que vous spécifiez peuvent voir une description localisée du complément dans AppSource (ou dans un catalogue privé).</span><span class="sxs-lookup"><span data-stu-id="f0124-128">Users in each locale you specify can see a localized description for the add-in in AppSource (or private catalog).</span></span><br/><span data-ttu-id="f0124-129">Pour les compléments Outlook, les utilisateurs peuvent voir le nom d’affichage sous forme d’étiquette pour le bouton de l’application Outlook ainsi que dans l’EAC après l’installation.</span><span class="sxs-lookup"><span data-stu-id="f0124-129">For Outlook add-ins, users can see the display name as a label for the Outlook add-in button and in the EAC after installation.</span></span><br/><span data-ttu-id="f0124-130">Pour les compléments de contenu et du volet Office, les utilisateurs peuvent voir l’icône dans le ruban après avoir installé l’application.</span><span class="sxs-lookup"><span data-stu-id="f0124-130">For content and task pane add-ins, users can see the display name in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="f0124-131">[IconUrl]</span><span class="sxs-lookup"><span data-stu-id="f0124-131">[IconUrl]</span></span>        |<span data-ttu-id="f0124-p105">L’image de l’icône est facultative. Vous pouvez utiliser la même technique de remplacement pour spécifier une image donnée pour une culture particulière. Si vous utilisez et localisez une icône, les utilisateurs de chaque paramètre régional que vous spécifiez peuvent voir l’image d’icône localisée pour le complément.</span><span class="sxs-lookup"><span data-stu-id="f0124-p105">The icon image is optional. You can use the same override technique to specify a certain image for a specific culture. If you use and localize an icon, users in each locale you specify can see a localized icon image for the add-in.</span></span><br/><span data-ttu-id="f0124-135">Pour les compléments Outlook, les utilisateurs peuvent voir l’icône dans l’EAC après l’installation du complément.</span><span class="sxs-lookup"><span data-stu-id="f0124-135">For Outlook add-ins, users can see the icon in the EAC after installing the add-in.</span></span><br/><span data-ttu-id="f0124-136">Pour les compléments de contenu et du volet de tâches, les utilisateurs peuvent voir l’icône dans le ruban après avoir installé le complément.</span><span class="sxs-lookup"><span data-stu-id="f0124-136">For content and task pane add-ins, users can see the icon in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="f0124-137">[HighResolutionIconUrl] **Important :** cet élément est disponible uniquement lors de l’utilisation de la version 1.1 du manifeste de complément.</span><span class="sxs-lookup"><span data-stu-id="f0124-137">[HighResolutionIconUrl] **Important:** This element is available only when using add-in manifest version 1.1.</span></span>|<span data-ttu-id="f0124-p106">L’image de l’icône de haute résolution est facultative. Néanmoins, si elle est indiquée, elle doit l’être après l’élément [IconUrl]. Si  [HighResolutionIconUrl] est spécifié et que le complément est installé sur un appareil qui prend en charge la haute résolution (dpi), la valeur [HighResolutionIconUrl] est utilisée à la place de la valeur [IconUrl].</span><span class="sxs-lookup"><span data-stu-id="f0124-p106">The high resolution icon image is optional but if it is specified, it must occur after the  [IconUrl] element. When [HighResolutionIconUrl] is specified, and the add-in is installed on a device that supports high dpi resolution, the [HighResolutionIconUrl] value is used instead of the value for [IconUrl].</span></span><br/><span data-ttu-id="f0124-p107">Si  HighResolutionIconUrl est spécifié et que le complément est installé sur un appareil qui prend en charge la haute résolution (dpi), la valeur HighResolutionIconUrl est utilisée à la place de la valeur IconUrl.</span><span class="sxs-lookup"><span data-stu-id="f0124-p107">You can use the same override technique to specify a certain image for a specific culture. If you use and localize an icon, users in each locale you specify can see a localized icon image for the add-in.</span></span><br/><span data-ttu-id="f0124-142">Si vous utilisez et localisez une icône, les utilisateurs de chaque paramètre régional que vous spécifiez peuvent voir l’image d’icône localisée pour le complément.</span><span class="sxs-lookup"><span data-stu-id="f0124-142">For Outlook add-ins, users can see the icon in the EAC after installing the add-in.</span></span><br/><span data-ttu-id="f0124-143">Pour les compléments Outlook, les utilisateurs peuvent voir l’icône dans l’EAC après l’installation du complément.</span><span class="sxs-lookup"><span data-stu-id="f0124-143">For content and task pane add-ins, users can see the icon in the ribbon after installing the add-in.</span></span>|
|<span data-ttu-id="f0124-144">Pour les compléments de contenu et du volet de tâches, les utilisateurs peuvent voir l’icône dans le ruban après avoir installé le complément.</span><span class="sxs-lookup"><span data-stu-id="f0124-144">[Resources] **Important:** This element is available only when using add-in manifest version 1.1.</span></span>   |<span data-ttu-id="f0124-145">[Ressources] Important : cet élément est disponible uniquement lors de l’utilisation de la version 1.1 du manifeste de complément.</span><span class="sxs-lookup"><span data-stu-id="f0124-145">Users in each locale you specify can see string and icon resources that you specifically create for the add-in for that locale.</span></span> |
|<span data-ttu-id="f0124-146">[SourceLocation]</span><span class="sxs-lookup"><span data-stu-id="f0124-146">[SourceLocation]</span></span>   |<span data-ttu-id="f0124-147">Les utilisateurs de chaque paramètre régional que vous spécifiez peuvent voir une page web que vous concevez spécifiquement pour le complément pour ce paramètre régional.</span><span class="sxs-lookup"><span data-stu-id="f0124-147">Users in each locale you specify can see a webpage that you specifically design for the add-in for that locale.</span></span> |


> [!NOTE]
> <span data-ttu-id="f0124-148">Vous pouvez localiser la description et le nom d’affichage uniquement pour les paramètres régionaux qu’Office prend en charge.</span><span class="sxs-lookup"><span data-stu-id="f0124-148">You can localize the description and display name for only the locales that Office supports.</span></span> <span data-ttu-id="f0124-149">Pour obtenir la liste des langues et les paramètres régionaux pour la version actuelle d’Office, voir [Identificateurs de langue et valeurs d’ID de l’élément OptionState dans Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15)).</span><span class="sxs-lookup"><span data-stu-id="f0124-149">See [Language identifiers and OptionState Id values in Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15)) for a list of languages and locales for the current release of Office.</span></span>


### <a name="examples"></a><span data-ttu-id="f0124-150">Exemples</span><span class="sxs-lookup"><span data-stu-id="f0124-150">Examples</span></span>

<span data-ttu-id="f0124-p109">Par exemple, un complément Office peut spécifier [DefaultLocale] en tant que `en-us`. Pour l’élément [DisplayName], le complément peut spécifier un élément enfant [Override] pour le paramètre régional `fr-fr`, comme illustré ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="f0124-p109">For example, an Office Add-in can specify the [DefaultLocale] as `en-us`. For the [DisplayName] element, the add-in can specify an [Override] child element for the locale `fr-fr`, as shown below.</span></span>


```xml
<DefaultLocale>en-us</DefaultLocale>
...
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

> [!NOTE]
> <span data-ttu-id="f0124-153">Si vous devez localiser plusieurs domaines au sein d’une famille de langues, comme `de-de` et `de-at`, nous vous recommandons d’utiliser des éléments `Override` distincts pour chaque domaine.</span><span class="sxs-lookup"><span data-stu-id="f0124-153">If you need to localize for more than one area within a language family, such as `de-de` and `de-at`, we recommend that you use separate `Override` elements for each area.</span></span> <span data-ttu-id="f0124-154">L’utilisation du seul nom de langue, dans ce cas, n’est pas prise en charge sur toutes les combinaisons d’applications et `de` de plateformes Office clientes.</span><span class="sxs-lookup"><span data-stu-id="f0124-154">Using just the language name alone, in this case, `de`, is not supported across all combinations of Office client applications and platforms.</span></span>

<span data-ttu-id="f0124-p111">Cela signifie que le complément adopte le paramètre régional `en-us` par défaut. Les utilisateurs voient le nom d’affichage « Video player » pour tous les paramètres régionaux, sauf si le paramètre régional de l’ordinateur client est `fr-fr`, auquel cas ils verront le nom d’affichage « Lecteur vidéo ».</span><span class="sxs-lookup"><span data-stu-id="f0124-p111">This means that the add-in assumes the  `en-us` locale by default. Users see the English display name of "Video player" for all locales unless the client computer's locale is `fr-fr`, in which case users would see the French display name "Lecteur vidéo".</span></span>

> [!NOTE]
> <span data-ttu-id="f0124-157">Vous ne pouvez spécifier qu’un seul remplacement par langue, notamment pour les paramètres régionaux par défaut.</span><span class="sxs-lookup"><span data-stu-id="f0124-157">You may only specify a single override per language, including for the default locale.</span></span> <span data-ttu-id="f0124-158">Par exemple, si votre paramètre régional par défaut est `en-us`, vous ne pouvez pas spécifier un remplacement pour `en-us`.</span><span class="sxs-lookup"><span data-stu-id="f0124-158">For example, if your default locale is `en-us` you cannot not specify an  override for `en-us` as well.</span></span> 

<span data-ttu-id="f0124-p113">L’exemple suivant applique un remplacement de paramètre régional pour l’élément [Description]. Il commence par spécifier le paramètre régional par défaut `en-us` et une description en anglais, puis spécifie une instruction [Override] avec une description en français pour le paramètre régional `fr-fr` :</span><span class="sxs-lookup"><span data-stu-id="f0124-p113">The following example applies a locale override for the [Description] element. It first specifies a default locale of `en-us` and an English description, and then specifies an [Override] statement with a French description for the `fr-fr` locale:</span></span>

```xml
<DefaultLocale>en-us</DefaultLocale>
...
<Description DefaultValue=
   "Watch YouTube videos referenced in the emails you receive
   without leaving your email client.">
   <Override Locale="fr-fr" Value=
   "Visualisez les vidéos YouTube référencées dans vos courriers 
   électronique directement depuis Outlook."/>
</Description>
```

<span data-ttu-id="f0124-p114">Il commence par spécifier le paramètre régional par défaut `en-us` et une description en anglais, puis spécifie une instruction `DefaultValue` avec une description en français pour le paramètre régional `fr-fr`:</span><span class="sxs-lookup"><span data-stu-id="f0124-p114">This means that the add-in assumes the `en-us` locale by default. Users would see the English description in the `DefaultValue` attribute for all locales unless the client computer's locale is `fr-fr`, in which case they would see the French description.</span></span>

<span data-ttu-id="f0124-p115">Les utilisateurs verront la description en anglais figurant dans l’attribut `fr-fr` pour tous les paramètres régionaux, sauf si le paramètre régional de l’ordinateur du client est `fr-fr`, auquel cas la description s’affichera en français.</span><span class="sxs-lookup"><span data-stu-id="f0124-p115">In the following example, the add-in specifies a separate image that's more appropriate for the `fr-fr` locale and culture. Users see the image DefaultLogo.png by default, except when the locale of the client computer is `fr-fr`. In this case, users would see the image FrenchLogo.png.</span></span> 


```xml
<!-- Replace "domain" with a real web server name and path. -->
<IconUrl DefaultValue="https://<domain>/DefaultLogo.png"/>
<Override Locale="fr-fr" Value="https://<domain>/FrenchLogo.png"/>
```

<span data-ttu-id="f0124-p116">Dans ce cas, les utilisateurs voient l’image FrenchLogo.png.</span><span class="sxs-lookup"><span data-stu-id="f0124-p116">The following example shows how to localize a resource in the `Resources` section. It applies a locale override for an image that is more appropriate for the `ja-jp` culture.</span></span>

```xml
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
 ...
```


<span data-ttu-id="f0124-p117">Une valeur de remplacement des paramètres régionaux est appliquée pour une image plus appropriée par rapport à la culture [].</span><span class="sxs-lookup"><span data-stu-id="f0124-p117">For the [SourceLocation] element, supporting additional locales means providing a separate source HTML file for each of the specified locales. Users in each locale you specify can see a customized webpage that you design for that them.</span></span>

<span data-ttu-id="f0124-p118">Les utilisateurs de chaque paramètre régional que vous spécifiez peuvent accéder à une page web personnalisée conçue pour eux.</span><span class="sxs-lookup"><span data-stu-id="f0124-p118">For Outlook add-ins, the [SourceLocation] element also aligns to the form factor. This allows you to provide a separate, localized source HTML file for each corresponding form factor. You can specify one or more [Override] child elements in each applicable settings element ([DesktopSettings], [TabletSettings], or [PhoneSettings]). The following example shows settings elements for the desktop, tablet, and smartphone form factors, each with one HTML file for the default locale and another for the French locale.</span></span>


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

## <a name="localize-extended-overrides"></a><span data-ttu-id="f0124-174">Localiser les substitutions étendues</span><span class="sxs-lookup"><span data-stu-id="f0124-174">Localize extended overrides</span></span>

<span data-ttu-id="f0124-175">Certaines fonctionnalités d’extensibilité des modules de Office, telles que les raccourcis clavier, sont configurées avec des fichiers JSON hébergés sur votre serveur, et non avec le manifeste XML du module.</span><span class="sxs-lookup"><span data-stu-id="f0124-175">Some extensibility features of Office Add-ins, such as keyboard shortcuts, are configured with JSON files that are hosted on your server, instead of with the add-in's XML manifest.</span></span> <span data-ttu-id="f0124-176">Cette section suppose que vous êtes familiarisé avec les substitutions étendues.</span><span class="sxs-lookup"><span data-stu-id="f0124-176">This section assumes that you're familiar with extended overrides.</span></span> <span data-ttu-id="f0124-177">Voir [Work with extended overrides of the manifest](extended-overrides.md) and [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span><span class="sxs-lookup"><span data-stu-id="f0124-177">See [Work with extended overrides of the manifest](extended-overrides.md) and [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span>

<span data-ttu-id="f0124-178">Utilisez `ResourceUrl` l’attribut de [l’élément ExtendedOverrides](../reference/manifest/extendedoverrides.md) pour pointer Office vers un fichier de ressources localisées.</span><span class="sxs-lookup"><span data-stu-id="f0124-178">Use the `ResourceUrl` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element to point Office to a file of localized resources.</span></span> <span data-ttu-id="f0124-179">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="f0124-179">The following is an example.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="f0124-180">Le fichier de remplacements étendu utilise ensuite des jetons au lieu de chaînes.</span><span class="sxs-lookup"><span data-stu-id="f0124-180">The extended overrides file then uses tokens instead of strings.</span></span> <span data-ttu-id="f0124-181">Chaînes de noms de jetons dans le fichier de ressources.</span><span class="sxs-lookup"><span data-stu-id="f0124-181">The tokens name strings in the resource file.</span></span> <span data-ttu-id="f0124-182">Voici un exemple qui affecte un raccourci clavier à une fonction (définie ailleurs) qui affiche le volet Des tâches du module.</span><span class="sxs-lookup"><span data-stu-id="f0124-182">The following is an example that assigns a keyboard shortcut to a function (defined elsewhere) that displays the add-in's task pane.</span></span> <span data-ttu-id="f0124-183">Remarque à propos de ce markup :</span><span class="sxs-lookup"><span data-stu-id="f0124-183">Note about this markup:</span></span>

- <span data-ttu-id="f0124-184">L’exemple n’est pas tout à fait valide.</span><span class="sxs-lookup"><span data-stu-id="f0124-184">The example isn't quite valid.</span></span> <span data-ttu-id="f0124-185">(Nous y ajoutons une propriété supplémentaire obligatoire ci-dessous.)</span><span class="sxs-lookup"><span data-stu-id="f0124-185">(We add a required additional property to it below.)</span></span>
- <span data-ttu-id="f0124-186">Les jetons doivent avoir le format **${resource.*nom de ressource*}**.</span><span class="sxs-lookup"><span data-stu-id="f0124-186">The tokens must have the format **${resource.*name-of-resource*}**.</span></span>

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "${resource.SHOWTASKPANE_action_name}"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "${resource.SHOWTASKPANE_default_shortcut}"
            }
        }
    ] 
}
```

<span data-ttu-id="f0124-187">Le fichier de ressources, également au format JSON, possède une propriété de niveau supérieur divisée en sous-propriétés par `resources` paramètres régionaux.</span><span class="sxs-lookup"><span data-stu-id="f0124-187">The resource file, which is also JSON-formatted, has a top-level `resources` property that is divided into subproperties by locale.</span></span> <span data-ttu-id="f0124-188">Pour chaque paramètre local, une chaîne est affectée à chaque jeton utilisé dans le fichier de remplacements étendu.</span><span class="sxs-lookup"><span data-stu-id="f0124-188">For each locale, a string is assigned to each token that was used in the extended overrides file.</span></span> <span data-ttu-id="f0124-189">Voici un exemple qui possède des chaînes pour `en-us` et `fr-fr` .</span><span class="sxs-lookup"><span data-stu-id="f0124-189">The following is an example which has strings for `en-us` and `fr-fr`.</span></span> <span data-ttu-id="f0124-190">Dans cet exemple, le raccourci clavier est le même dans les deux paramètres régionaux, mais ce n’est pas toujours le cas, en particulier lorsque vous localisez des paramètres régionaux dont l’alphabet ou le système d’écriture est différent, et par conséquent un autre clavier.</span><span class="sxs-lookup"><span data-stu-id="f0124-190">In this example, the keyboard shortcut is the same in both locales, but that won't always be the case, especially when you are localizing for locales that have a different alphabet or writing system, and hence a different keyboard.</span></span>

```json
{
    "resources":{ 
        "en-us": { 
            "SHOWTASKPANE_default_shortcut": { 
                "value": "CTRL+SHIFT+A", 
            }, 
            "SHOWTASKPANE_action_name": {
                "value": "Show task pane for add-in",
            }, 
        },
        "fr-fr": { 
            "SHOWTASKPANE_default_shortcut": { 
                "value": "CTRL+SHIFT+A", 
            }, 
            "SHOWTASKPANE_action_name": {
                "value": "Afficher le volet de tâche pour add-in",
              } 
        }
    }
}
```

<span data-ttu-id="f0124-191">Il `default` n’existe aucune propriété dans le fichier qui soit un homologue aux `en-us` `fr-fr` sections et aux sections.</span><span class="sxs-lookup"><span data-stu-id="f0124-191">There is no `default` property in the file that is a peer to the `en-us` and `fr-fr` sections.</span></span> <span data-ttu-id="f0124-192">En effet, les chaînes par défaut, qui sont utilisées lorsque les paramètres régionaux de l’application hôte Office ne correspondent à aucune des propriétés *ll-cc* dans le fichier de ressources, doivent être définies dans le fichier de remplacements étendu *lui-même.*</span><span class="sxs-lookup"><span data-stu-id="f0124-192">This is because the default strings, which are used when the locale of the Office host application doesn't match any of the *ll-cc* properties in the resources file, *must be defined in the extended overrides file itself*.</span></span> <span data-ttu-id="f0124-193">La définition des chaînes par défaut directement dans le fichier de remplacements étendu garantit que Office ne télécharge pas le fichier de ressources lorsque les paramètres régionaux de l’application Office sont les paramètres régionaux par défaut du module (comme spécifié dans le manifeste).</span><span class="sxs-lookup"><span data-stu-id="f0124-193">Defining the default strings directly in the extended overrides file ensures that Office doesn't download the resource file when the locale of the Office application matches the default locale of the add-in (as specified in the manifest).</span></span> <span data-ttu-id="f0124-194">Voici une version corrigée de l’exemple précédent d’un fichier de remplacements étendu qui utilise des jetons de ressource.</span><span class="sxs-lookup"><span data-stu-id="f0124-194">The following is a corrected version of the preceding example of an extended overrides file that uses resource tokens.</span></span>

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "${resource.SHOWTASKPANE_action_name}"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "${resource.SHOWTASKPANE_default_shortcut}"
            }
        }
    ],
    "resources": { 
        "default": { 
            "SHOWTASKPANE_default_shortcut": { 
                "value": "CTRL+SHIFT+A", 
            }, 
            "SHOWTASKPANE_action_name": {
                "value": "Show task pane for add-in",
            } 
        }
    }
}
```

## <a name="match-datetime-format-with-client-locale"></a><span data-ttu-id="f0124-195">Mettre en correspondance le format de date/heure avec le paramètre régional du client</span><span class="sxs-lookup"><span data-stu-id="f0124-195">Match date/time format with client locale</span></span>

<span data-ttu-id="f0124-196">Vous pouvez obtenir les paramètres régionaux de l’interface utilisateur de l’application Office client à l’aide de la **[propriété displayLanguage.]**</span><span class="sxs-lookup"><span data-stu-id="f0124-196">You can get the locale of the user interface of the Office client application by using the **[displayLanguage]** property.</span></span> <span data-ttu-id="f0124-197">Vous pouvez ensuite afficher les valeurs de date et d’heure dans un format cohérent avec les paramètres régionaux actuels de l Office application.</span><span class="sxs-lookup"><span data-stu-id="f0124-197">You can then display date and time values in a format consistent with the current locale of the Office application.</span></span> <span data-ttu-id="f0124-198">Vous pouvez ensuite afficher les valeurs de date et d’heure dans un format cohérent avec les paramètres régionaux actuels de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="f0124-198">One way to do that is to prepare a resource file that specifies the date/time display format to use for each locale that your Office Add-in supports.</span></span> <span data-ttu-id="f0124-199">Au moment de l’utilisation, votre add-in peut utiliser le fichier de ressources et faire correspondre le format de date/heure approprié aux paramètres régionaux obtenus à partir de la **[propriété displayLanguage.]**</span><span class="sxs-lookup"><span data-stu-id="f0124-199">At run time, your add-in can use the resource file and match the appropriate date/time format with the locale obtained from the **[displayLanguage]** property.</span></span>

<span data-ttu-id="f0124-200">Vous pouvez obtenir les paramètres régionaux des données de l’application Office client à l’aide de la [propriété contentLanguage.]</span><span class="sxs-lookup"><span data-stu-id="f0124-200">You can get the locale of the data of the Office client application by using the [contentLanguage] property.</span></span> <span data-ttu-id="f0124-201">Vous pouvez obtenir les paramètres régionaux des données de l’application d’hébergement en utilisant la propriété contentLanguage.</span><span class="sxs-lookup"><span data-stu-id="f0124-201">Based on this value, you can then appropriately interpret or display date/time strings.</span></span> <span data-ttu-id="f0124-202">En fonction de cette valeur, vous pouvez correctement interpréter ou afficher des chaînes de date/heure.</span><span class="sxs-lookup"><span data-stu-id="f0124-202">For example, the `jp-JP` locale expresses data/time values as `yyyy/MM/dd`, and the `fr-FR` locale, `dd/MM/yyyy`.</span></span>


## <a name="use-ajax-for-globalization-and-localization"></a><span data-ttu-id="f0124-203">Utiliser Ajax pour l’internationalisation et la localisation</span><span class="sxs-lookup"><span data-stu-id="f0124-203">Use Ajax for globalization and localization</span></span>


<span data-ttu-id="f0124-204">Si vous utilisez Visual Studio pour créer des Compléments Office, .NET Framework et Ajax offrent des moyens d’internationaliser et de localiser les fichiers de script client.</span><span class="sxs-lookup"><span data-stu-id="f0124-204">If you use Visual Studio to create Office Add-ins, the .NET Framework and Ajax provide ways to globalize and localize client script files.</span></span>

<span data-ttu-id="f0124-p127">Si vous utilisez Visual Studio pour créer des Compléments Office, .NET Framework et Ajax offrent des moyens d’internationaliser et de localiser les fichiers de script client.</span><span class="sxs-lookup"><span data-stu-id="f0124-p127">You can globalize and use the [Date](/previous-versions/bb310850(v=vs.140)) and [Number](/previous-versions/bb310835(v=vs.140)) JavaScript type extensions and the JavaScript [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object in the JavaScript code for an Office Add-in to display values based on the locale settings on the current browser. For more information, see [Walkthrough: Globalizing a Date by Using Client Script](/previous-versions/bb386581(v=vs.140)).</span></span>

<span data-ttu-id="f0124-p128">Pour plus d’informations, voir Walkthrough: Globalizing a Date by Using Client Script.</span><span class="sxs-lookup"><span data-stu-id="f0124-p128">You can include localized resource strings directly in standalone JavaScript files to provide client script files for different locales, which are set on the browser or provided by the user. Create a separate script file for each supported locale. In each script file, include an object in JSON format that contains the resource strings for that locale. The localized values are applied when the script runs in the browser.</span></span>


## <a name="example-build-a-localized-office-add-in"></a><span data-ttu-id="f0124-211">Exemple : créer un complément Office localisé</span><span class="sxs-lookup"><span data-stu-id="f0124-211">Example: Build a localized Office Add-in</span></span>

<span data-ttu-id="f0124-212">Cette section inclut des exemples expliquant comment localiser la description, le nom d’affichage et l’interface utilisateur d’une Complément Office.</span><span class="sxs-lookup"><span data-stu-id="f0124-212">This section provides examples that show you how to localize an Office Add-in description, display name, and UI.</span></span> 

> [!NOTE]
> <span data-ttu-id="f0124-213">Pour télécharger Visual Studio 2019, consultez la [page Visual Studio IDE.](https://visualstudio.microsoft.com/vs/)</span><span class="sxs-lookup"><span data-stu-id="f0124-213">To download Visual Studio 2019, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/).</span></span> <span data-ttu-id="f0124-214">Lors de l’installation, vous devez sélectionner la charge de travail de développement Office/SharePoint.</span><span class="sxs-lookup"><span data-stu-id="f0124-214">During installation you'll need to select the Office/SharePoint development workload.</span></span>

### <a name="configure-office-to-use-additional-languages-for-display-or-editing"></a><span data-ttu-id="f0124-215">Configurer Office pour utiliser des langues supplémentaires pour l’affichage ou l’édition</span><span class="sxs-lookup"><span data-stu-id="f0124-215">Configure Office to use additional languages for display or editing</span></span>

<span data-ttu-id="f0124-216">Pour exécuter l’exemple de code fourni, configurez Office sur votre ordinateur pour utiliser des langues supplémentaires afin de pouvoir tester votre complément en changeant la langue utilisée pour l’affichage dans les menus et les commandes, pour la modification et la mise en preuve, ou les deux.</span><span class="sxs-lookup"><span data-stu-id="f0124-216">To run the sample code provided, configure Office on your computer to use additional languages so that you can test your add-in by switching the language used for display in menus and commands, for editing and proofing, or both.</span></span>

<span data-ttu-id="f0124-217">Vous pouvez utiliser un module linguistique Office pour installer une autre langue.</span><span class="sxs-lookup"><span data-stu-id="f0124-217">You can use an Office Language pack to install an additional language.</span></span> <span data-ttu-id="f0124-218">Pour plus d’informations sur les Modules linguistiques et où les obtenir, voir [Pack d’accessoires linguistiques pour Office](https://support.microsoft.com/office/82ee1236-0f9a-45ee-9c72-05b026ee809f).</span><span class="sxs-lookup"><span data-stu-id="f0124-218">For more information about Language Packs and where to get them, see [Language Accessory Pack for Office](https://support.microsoft.com/office/82ee1236-0f9a-45ee-9c72-05b026ee809f).</span></span>

<span data-ttu-id="f0124-219">Après avoir installé le Pack d’accessoires linguistiques, vous pouvez configurer Office pour utiliser la langue installée pour l’affichage dans l’interface utilisateur, pour modifier du contenu de document, ou les deux.</span><span class="sxs-lookup"><span data-stu-id="f0124-219">After you install the Language Accessory Pack, you can configure Office to use the installed language for display in the UI, for editing document content, or both.</span></span> <span data-ttu-id="f0124-220">L’exemple dans cet article utilise une installation d’Office qui contient le module linguistique espagnol.</span><span class="sxs-lookup"><span data-stu-id="f0124-220">The example in this article uses an installation of Office that has the Spanish Language Pack applied.</span></span>

### <a name="create-an-office-add-in-project"></a><span data-ttu-id="f0124-221">Créer un projet de complément Office</span><span class="sxs-lookup"><span data-stu-id="f0124-221">Create an Office Add-in project</span></span>

<span data-ttu-id="f0124-222">Vous devez créer un projet de Visual Studio 2019 Office de recherche.</span><span class="sxs-lookup"><span data-stu-id="f0124-222">You'll need to create a Visual Studio 2019 Office Add-in project.</span></span>

> [!NOTE]
> <span data-ttu-id="f0124-223">Si vous n’avez pas installé Visual Studio 2019, consultez la [page Visual Studio IDE](https://visualstudio.microsoft.com/vs/) pour obtenir des instructions de téléchargement.</span><span class="sxs-lookup"><span data-stu-id="f0124-223">If you haven't installed Visual Studio 2019, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/) for download instructions.</span></span> <span data-ttu-id="f0124-224">Lors de l’installation, vous devez sélectionner la charge de travail de développement Office/SharePoint.</span><span class="sxs-lookup"><span data-stu-id="f0124-224">During installation you'll need to select the Office/SharePoint development workload.</span></span> <span data-ttu-id="f0124-225">Si vous avez déjà installé Visual Studio 2019, [](/visualstudio/install/modify-visual-studio/) utilisez la Visual Studio Installer pour vous assurer que la charge de travail de développement Office/SharePoint est installée.</span><span class="sxs-lookup"><span data-stu-id="f0124-225">If you have previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio/) to ensure that the Office/SharePoint development workload is installed.</span></span>

1. <span data-ttu-id="f0124-226">Choisissez **Créer un nouveau projet**.</span><span class="sxs-lookup"><span data-stu-id="f0124-226">Choose **Create a new project**.</span></span>

2. <span data-ttu-id="f0124-227">À l’aide de la zone de recherche, entrez **complément**.</span><span class="sxs-lookup"><span data-stu-id="f0124-227">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="f0124-228">Choisissez **Complément web Word**, puis sélectionnez **Suivant**.</span><span class="sxs-lookup"><span data-stu-id="f0124-228">Choose **Word Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="f0124-229">Nommez votre projet **WorldReadyAddIn** et sélectionnez **Créer.**</span><span class="sxs-lookup"><span data-stu-id="f0124-229">Name your project **WorldReadyAddIn** and select **Create**.</span></span>

4. <span data-ttu-id="f0124-p134">Visual Studio crée une solution et ses deux projets apparaissent dans l’**explorateur de solutions**. Le fichier **Home.html** s’ouvre dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="f0124-p134">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>


### <a name="localize-the-text-used-in-your-add-in"></a><span data-ttu-id="f0124-232">Localiser le texte utilisé dans votre complément</span><span class="sxs-lookup"><span data-stu-id="f0124-232">Localize the text used in your add-in</span></span>

<span data-ttu-id="f0124-233">Le texte que vous souhaitez localiser dans une autre langue apparaît à deux emplacements :</span><span class="sxs-lookup"><span data-stu-id="f0124-233">The text that you want to localize for another language appears in two areas:</span></span>

-  <span data-ttu-id="f0124-p135">**Nom d’affichage et description du complément**. Ce contenu est contrôlé par les entrées du fichier manifeste de l’application.</span><span class="sxs-lookup"><span data-stu-id="f0124-p135">**Add-in display name and description**. This is controlled by entries in the add-in manifest file.</span></span>

-  <span data-ttu-id="f0124-236">**Interface utilisateur du complément**.</span><span class="sxs-lookup"><span data-stu-id="f0124-236">**Add-in UI**.</span></span> <span data-ttu-id="f0124-237">Vous pouvez localiser les chaînes qui s’affichent dans l’interface utilisateur de votre complément à l’aide du code JavaScript, par exemple en utilisant un fichier de ressources séparé contenant les chaînes localisées.</span><span class="sxs-lookup"><span data-stu-id="f0124-237">You can localize the strings that appear in your add-in UI by using JavaScript code, for example, by using a separate resource file that contains the localized strings.</span></span>

<span data-ttu-id="f0124-238">Pour localiser le nom d’affichage et la description du complément</span><span class="sxs-lookup"><span data-stu-id="f0124-238">To localize the add-in display name and description:</span></span>

1. <span data-ttu-id="f0124-239">Dans l’**Explorateur de solutions**, développez **WorldReadyAddIn**, **WorldReadyAddInManifest**, puis choisissez **WorldReadyAddIn.xml**.</span><span class="sxs-lookup"><span data-stu-id="f0124-239">In **Solution Explorer**, expand **WorldReadyAddIn**, **WorldReadyAddInManifest**, and then choose **WorldReadyAddIn.xml**.</span></span>

2. <span data-ttu-id="f0124-240">Dans WorldReadyAddInManifest.xml, remplacez les éléments [DisplayName] et [Description] par le bloc de code suivant.</span><span class="sxs-lookup"><span data-stu-id="f0124-240">In WorldReadyAddInManifest.xml, replace the [DisplayName] and [Description] elements with the following block of code.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f0124-241">Vous pouvez remplacer les chaînes localisées en espagnol utilisées dans cet exemple pour les éléments [DisplayName] et [Description] par les chaînes localisées en une autre langue.</span><span class="sxs-lookup"><span data-stu-id="f0124-241">You can replace the Spanish language localized strings used in this example for the [DisplayName] and [Description] elements with the localized strings for any other language.</span></span>

    ```xml
    <DisplayName DefaultValue="World Ready add-in">
      <Override Locale="es-es" Value="Aplicación de uso internacional"/>
    </DisplayName>
    <Description DefaultValue="An add-in for testing localization">
      <Override Locale="es-es" Value="Una aplicación para la prueba de la localización"/>
    </Description>
    ```

3. <span data-ttu-id="f0124-242">Lorsque vous modifiez la langue d’affichage dans Office 2013, par exemple de l’anglais vers l’espagnol, puis que vous exécutez le complément, le nom d’affichage et la description du complément sont affichés avec le texte localisé.</span><span class="sxs-lookup"><span data-stu-id="f0124-242">When you change the display language for Office 2013 from English to Spanish, for example, and then run the add-in, the add-in display name and description are shown with localized text.</span></span>

<span data-ttu-id="f0124-243">Lorsque vous modifiez la langue d’affichage dans Office 2013, par exemple de l’anglais vers l’espagnol, puis que vous exécutez le complément, le nom d’affichage et la description du complément sont affichés avec le texte localisé.</span><span class="sxs-lookup"><span data-stu-id="f0124-243">To lay out the add-in UI:</span></span>

1. <span data-ttu-id="f0124-244">Dans Visual Studio, dans l’**Explorateur de solutions**, choisissez **Home.html**.</span><span class="sxs-lookup"><span data-stu-id="f0124-244">In Visual Studio, in **Solution Explorer**, choose **Home.html**.</span></span>

2. <span data-ttu-id="f0124-245">Remplacez le contenu de l’élément `<body>` dans Home.html par le HTML suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="f0124-245">Replace the `<body>` element contents in Home.html with the following HTML, and save the file.</span></span>

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

<span data-ttu-id="f0124-246">La figure suivante montre l’élément titre (h1) et l’élément paragraphe (p) qui afficheront le texte localisé lorsque vous terminez les étapes restantes et exécutez le complément.</span><span class="sxs-lookup"><span data-stu-id="f0124-246">The following figure shows the heading (h1) element and the paragraph (p) element that will display localized text when you complete the remaining steps and run the add-in.</span></span>

<span data-ttu-id="f0124-247">*Figure 1. Interface utilisateur du complément*</span><span class="sxs-lookup"><span data-stu-id="f0124-247">*Figure 1. The add-in UI*</span></span>

![Interface utilisateur de l’application avec des sections en surbrillance](../images/office15-app-how-to-localize-fig03.png)

### <a name="add-the-resource-file-that-contains-the-localized-strings"></a><span data-ttu-id="f0124-249">Ajouter le fichier de ressources qui contient les chaînes localisées</span><span class="sxs-lookup"><span data-stu-id="f0124-249">Add the resource file that contains the localized strings</span></span>

<span data-ttu-id="f0124-250">Le fichier de ressource JavaScript contient les chaînes utilisées pour l’interface utilisateur du complément.</span><span class="sxs-lookup"><span data-stu-id="f0124-250">The JavaScript resource file contains the strings used for the add-in UI.</span></span> <span data-ttu-id="f0124-251">Le code HTML pour l’exemple d’interface utilisateur du complément contient un `<h1>` élément qui affiche un message d’accueil et un `<p>` élément qui présente le complément à l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f0124-251">The HTML for the sample add-in UI contains an `<h1>` element that displays a greeting, and a `<p>` element that introduces the add-in to the user.</span></span> 

<span data-ttu-id="f0124-p138">Pour activer les chaînes localisées pour le titre et le paragraphe, placez les chaînes dans un fichier de ressources distinct. Le fichier de ressources crée un objet JavaScript qui contient un objet JavaScript Object Notation (JSON) individuel pour chaque ensemble de chaînes localisées. Le fichier de ressources fournit une méthode pour obtenir l’objet JSON approprié pour des paramètres régionaux donnés.</span><span class="sxs-lookup"><span data-stu-id="f0124-p138">To enable localized strings for the heading and paragraph, you place the strings in a separate resource file. The resource file creates a JavaScript object that contains a separate JavaScript Object Notation (JSON) object for each set of localized strings. The resource file also provides a method for getting back the appropriate JSON object for a given locale.</span></span>

<span data-ttu-id="f0124-255">Pour ajouter le fichier de ressources au projet de complément :</span><span class="sxs-lookup"><span data-stu-id="f0124-255">To add the resource file to the add-in project:</span></span>

1. <span data-ttu-id="f0124-256">Dans **l’Explorateur de solutions** dans Visual Studio, cliquez avec le bouton droit sur le projet **WorldReadyAddInWeb**, puis choisissez **Ajouter** > **Nouvel élément**.</span><span class="sxs-lookup"><span data-stu-id="f0124-256">In **Solution Explorer** in Visual Studio, right-click the **WorldReadyAddInWeb** project and choose **Add** > **New Item**.</span></span> 

2. <span data-ttu-id="f0124-257">Dans la boîte de dialogue **Ajouter un nouvel élément**, choisissez **Fichier JavaScript**.</span><span class="sxs-lookup"><span data-stu-id="f0124-257">In the **Add New Item** dialog box, choose **JavaScript File**.</span></span>

3. <span data-ttu-id="f0124-258">Entrez **UIStrings.js** comme nom de fichier puis sélectionnez **Ajouter**.</span><span class="sxs-lookup"><span data-stu-id="f0124-258">Enter **UIStrings.js** as the file name and choose **Add**.</span></span>

4. <span data-ttu-id="f0124-259">Ajoutez le code suivant au fichier UIStrings.js et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="f0124-259">Add the following code to the UIStrings.js file, and save the file.</span></span>

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

<span data-ttu-id="f0124-260">Le fichier de ressources UIStrings.js crée un objet **UIStrings** qui contient les chaînes localisées pour l’interface utilisateur de votre complément.</span><span class="sxs-lookup"><span data-stu-id="f0124-260">The UIStrings.js resource file creates an object, **UIStrings**, which contains the localized strings for your add-in UI.</span></span>

### <a name="localize-the-text-used-for-the-add-in-ui"></a><span data-ttu-id="f0124-261">Localiser le texte utilisé pour l’interface utilisateur du complément</span><span class="sxs-lookup"><span data-stu-id="f0124-261">Localize the text used for the add-in UI</span></span>

<span data-ttu-id="f0124-p139">Pour utiliser le fichier de ressources de votre complément, vous devez ajouter une balise de script pour ce fichier dans Home.html. Quand Home.html est chargé, UIStrings.js s’exécute et l’objet  **UIStrings** que vous utilisez pour obtenir les chaînes est disponible pour votre code. Ajoutez le code HTML suivant dans la balise head pour Home.html pour que **UIStrings** soit disponible pour votre code.</span><span class="sxs-lookup"><span data-stu-id="f0124-p139">To use the resource file in your add-in, you'll need to add a script tag for it on Home.html. When Home.html is loaded, UIStrings.js executes and the **UIStrings** object that you use to get the strings is available to your code. Add the following HTML in the head tag for Home.html to make **UIStrings** available to your code.</span></span>

```html
<!-- Resource file for localized strings: -->
<script src="../UIStrings.js" type="text/javascript"></script>
```

<span data-ttu-id="f0124-265">Ajoutez le code HTML suivant dans la balise head pour Home.html pour que **UIStrings** soit disponible pour votre code.</span><span class="sxs-lookup"><span data-stu-id="f0124-265">Now you can use the **UIStrings** object to set the strings for the UI of your add-in.</span></span>

<span data-ttu-id="f0124-266">Si vous souhaitez modifier la localisation de votre application en fonction de la langue utilisée pour l’affichage dans les menus et les commandes de l’application cliente Office, utilisez la propriété **Office.context.displayLanguage** pour obtenir les paramètres régionaux de cette langue.</span><span class="sxs-lookup"><span data-stu-id="f0124-266">If you want to change the localization for your add-in based on what language is used for display in menus and commands in the Office client application, you use the **Office.context.displayLanguage** property to get the locale for that language.</span></span> <span data-ttu-id="f0124-267">Par exemple, si la langue de l’application utilise l’espagnol pour l’affichage dans les menus et les commandes, la propriété **Office.context.displayLanguage** retourne le code de langue es-ES.</span><span class="sxs-lookup"><span data-stu-id="f0124-267">For example, if the application language uses Spanish for display in menus and commands, the **Office.context.displayLanguage** property will return the language code es-ES.</span></span>

<span data-ttu-id="f0124-268">Si vous souhaitez modifier la localisation de votre add-in en fonction de la langue utilisée pour modifier le contenu du document, utilisez la propriété **Office.context.contentLanguage** pour obtenir les paramètres régionaux de cette langue.</span><span class="sxs-lookup"><span data-stu-id="f0124-268">If you want to change the localization for your add-in based on what language is being used for editing document content, you use the **Office.context.contentLanguage** property to get the locale for that language.</span></span> <span data-ttu-id="f0124-269">Par exemple, si la langue de l’application utilise l’espagnol pour modifier le contenu du document, la propriété **Office.context.contentLanguage** retourne le code de langue es-ES.</span><span class="sxs-lookup"><span data-stu-id="f0124-269">For example, if the application language uses Spanish for editing document content, the **Office.context.contentLanguage** property will return the language code es-ES.</span></span>

<span data-ttu-id="f0124-270">Une fois que vous connaissez la langue que l’application utilise, vous pouvez utiliser **UIStrings** pour obtenir l’ensemble de chaînes localisées qui correspond à la langue de l’application.</span><span class="sxs-lookup"><span data-stu-id="f0124-270">After you know the language the application is using, you can use **UIStrings** to get the set of localized strings that matches the application language.</span></span>

<span data-ttu-id="f0124-271">Remplacez le code du fichier Home.js par le code suivant. Le code montre comment changer les chaînes utilisées dans les éléments d’interface utilisateur de Home.html en fonction de la langue d’affichage de l’application hôte ou de la langue d’édition de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="f0124-271">Replace the code in the Home.js file with the following code.</span></span> <span data-ttu-id="f0124-272">Le code montre comment vous pouvez modifier les chaînes utilisées dans les éléments d’interface utilisateur sur Home.html en fonction de la langue d’affichage de l’application ou de la langue d’édition de l’application.</span><span class="sxs-lookup"><span data-stu-id="f0124-272">The code shows how you can change the strings used in the UI elements on Home.html based on either the display language of the application or the editing language of the application.</span></span>

> [!NOTE]
> <span data-ttu-id="f0124-273">Pour activer ou désactiver la localisation du complément en fonction de la langue utilisée pour l’édition, supprimez le commentaire de la ligne de code `var myLanguage = Office.context.contentLanguage;` et ajoutez un commentaire à la ligne de code `var myLanguage = Office.context.displayLanguage;`</span><span class="sxs-lookup"><span data-stu-id="f0124-273">To switch between changing the localization of the add-in based on the language used for editing, uncomment the line of code  `var myLanguage = Office.context.contentLanguage;` and comment out the line of code `var myLanguage = Office.context.displayLanguage;`</span></span>

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

            // Get the language setting for UI display in the Office application.
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

### <a name="test-your-localized-add-in"></a><span data-ttu-id="f0124-274">Tester votre complément localisé</span><span class="sxs-lookup"><span data-stu-id="f0124-274">Test your localized add-in</span></span>

<span data-ttu-id="f0124-275">Pour tester votre application localisée, modifiez la langue utilisée pour l’affichage ou la modification dans l’application Office puis exécutez votre application.</span><span class="sxs-lookup"><span data-stu-id="f0124-275">To test your localized add-in, change the language used for display or editing in the Office application and then run your add-in.</span></span>

<span data-ttu-id="f0124-276">Pour changer la langue utilisée pour l’affichage ou l’édition dans votre complément :</span><span class="sxs-lookup"><span data-stu-id="f0124-276">To change the language used for display or editing in your add-in:</span></span>

1. <span data-ttu-id="f0124-277">Dans Word, sélectionnez **Fichier** > **Options** > **Langue**.</span><span class="sxs-lookup"><span data-stu-id="f0124-277">In Word, choose **File** > **Options** > **Language**.</span></span> <span data-ttu-id="f0124-278">La figure suivante montre la boîte de dialogue **Options Word** ouverte sous l’onglet Langue.</span><span class="sxs-lookup"><span data-stu-id="f0124-278">The following figure shows the **Word Options** dialog box opened to the Language tab.</span></span>

    <span data-ttu-id="f0124-279">*Figure 2. Options de langue dans la boîte de dialogue Options Word*</span><span class="sxs-lookup"><span data-stu-id="f0124-279">*Figure 2. Language options in the Word Options dialog box*</span></span>

    ![Boîte de dialogue Options Word.](../images/office15-app-how-to-localize-fig04.png)

2. <span data-ttu-id="f0124-281">Sous **Choisir la langue d’affichage**, sélectionnez la langue que vous souhaitez afficher, par exemple espagnol, puis sélectionnez la flèche vers le haut pour déplacer la langue Espagnol en première position dans la liste.</span><span class="sxs-lookup"><span data-stu-id="f0124-281">Under **Choose Display Language**, select the language that you want for display, for example Spanish, and then choose the up arrow to move the Spanish language to the first position in the list.</span></span> <span data-ttu-id="f0124-282">Vous pouvez également modifier la langue utilisée pour la modification, sous Choisir les **langues** d’édition, choisissez la langue que vous souhaitez utiliser pour la modification, par exemple, l’espagnol, puis choisissez Définir par **défaut.**</span><span class="sxs-lookup"><span data-stu-id="f0124-282">Alternatively, to change the language used for editing, under **Choose Editing Languages**, choose the language you want to use for editing, for example, Spanish, and then choose **Set as Default**.</span></span>

3. <span data-ttu-id="f0124-283">Sélectionnez **OK** pour confirmer votre choix, puis fermez Word.</span><span class="sxs-lookup"><span data-stu-id="f0124-283">Choose **OK** to confirm your selection, and then close Word.</span></span>

4. <span data-ttu-id="f0124-284">Appuyez sur **F5** dans Visual Studio pour exécuter le complément d’exemple ou choisissez **Déboguer** > **Démarrer le débogage** dans la barre de menus.</span><span class="sxs-lookup"><span data-stu-id="f0124-284">Press **F5** in Visual Studio to run the sample add-in, or choose **Debug** > **Start Debugging** from the menu bar.</span></span>

5. <span data-ttu-id="f0124-285">Dans Word, sélectionnez **Accueil** > **Afficher le volet de tâches**.</span><span class="sxs-lookup"><span data-stu-id="f0124-285">In Word, choose **Home** > **Show Taskpane**.</span></span>

<span data-ttu-id="f0124-286">Une fois l’exécution en cours d’exécution, les chaînes de l’interface utilisateur du add-in changent pour correspondre à la langue utilisée par l’application, comme illustré dans la figure suivante.</span><span class="sxs-lookup"><span data-stu-id="f0124-286">Once running, the strings in the add-in UI change to match the language used by the application, as shown in the following figure.</span></span>


<span data-ttu-id="f0124-287">Le complément de volet de tâches est chargé dans Word 2013 et les chaînes de l’interface utilisateur du complément changent pour correspondre à la langue utilisée par l’application hôte, comme indiqué dans la figure suivante.</span><span class="sxs-lookup"><span data-stu-id="f0124-287">*Figure 3. Add-in UI with localized text*</span></span>

![Application avec texte de l’interface utilisateur localisé](../images/office15-app-how-to-localize-fig05.png)

## <a name="see-also"></a><span data-ttu-id="f0124-289">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f0124-289">See also</span></span>

- [<span data-ttu-id="f0124-290">Instructions de conception pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="f0124-290">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- <span data-ttu-id="f0124-291">[Instructions de conception pour les compléments Office](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))</span><span class="sxs-lookup"><span data-stu-id="f0124-291">[Language identifiers and OptionState Id values in Office 2013](/previous-versions/office/office-2013-resource-kit/cc179219(v=office.15))</span></span>

[DefaultLocale]:         ../reference/manifest/defaultlocale.md
[Description]:           ../reference/manifest/description.md
[DisplayName]:           ../reference/manifest/displayname.md
[IconUrl]:               ../reference/manifest/iconurl.md
[HighResolutionIconUrl]: ../reference/manifest/highresolutioniconurl.md
[Resources]:             ../reference/manifest/resources.md
[SourceLocation]:        ../reference/manifest/sourcelocation.md
[Override]:              ../reference/manifest/override.md
[DesktopSettings]:       ../reference/manifest/desktopsettings.md
[TabletSettings]:        ../reference/manifest/tabletsettings.md
[PhoneSettings]:         ../reference/manifest/phonesettings.md
[displayLanguage]:       /javascript/api/office/office.context#displaylanguage
[contentLanguage]:       /javascript/api/office/office.context#contentlanguage
[RFC 3066]:              https://www.rfc-editor.org/info/rfc3066
[RFC 3066]:              https://www.rfc-editor.org/info/rfc3066
