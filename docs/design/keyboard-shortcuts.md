---
title: Raccourcis clavier personnalisés dans les compléments Office
description: Découvrez comment ajouter des raccourcis clavier personnalisés, également appelés combinaisons de touches, à votre complément Office.
ms.date: 11/09/2020
localization_priority: Normal
ms.openlocfilehash: 40009dd92787b7c220bb8cfc741cffb2e4b68a9e
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132038"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a><span data-ttu-id="8e830-103">Ajouter des raccourcis clavier personnalisés à vos compléments Office (aperçu)</span><span class="sxs-lookup"><span data-stu-id="8e830-103">Add Custom keyboard shortcuts to your Office Add-ins (preview)</span></span>

<span data-ttu-id="8e830-104">Les raccourcis clavier, également appelés combinaisons de touches, permettent aux utilisateurs de votre complément de travailler plus efficacement et améliorent l’accessibilité du complément pour les utilisateurs présentant un handicap en fournissant une alternative à la souris.</span><span class="sxs-lookup"><span data-stu-id="8e830-104">Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently and they improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.</span></span>

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> <span data-ttu-id="8e830-105">Pour commencer avec une version de travail d’un complément avec des raccourcis clavier déjà activés, clonez et exécutez l’exemple de [raccourcis clavier Excel](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span><span class="sxs-lookup"><span data-stu-id="8e830-105">To start with a working version of an add-in with keyboard shortcuts already enabled, clone and run the sample [Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span> <span data-ttu-id="8e830-106">Lorsque vous êtes prêt à ajouter des raccourcis clavier à votre propre complément, poursuivez avec cet article.</span><span class="sxs-lookup"><span data-stu-id="8e830-106">When you are ready to add keyboard shortcuts to your own add-in, continue with this article.</span></span>

<span data-ttu-id="8e830-107">Il y a trois étapes pour ajouter des raccourcis clavier à un complément :</span><span class="sxs-lookup"><span data-stu-id="8e830-107">There are three steps to add keyboard shortcuts to an add-in:</span></span>

1. <span data-ttu-id="8e830-108">[Configurez le manifeste du complément](#configure-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="8e830-108">[Configure the add-in's manifest](#configure-the-manifest).</span></span>
1. <span data-ttu-id="8e830-109">[Créez ou modifiez le fichier JSON de raccourcis](#create-or-edit-the-shortcuts-json-file) pour définir les actions et leurs raccourcis clavier.</span><span class="sxs-lookup"><span data-stu-id="8e830-109">[Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file) to define actions and their keyboard shortcuts.</span></span>
1. <span data-ttu-id="8e830-110">[Ajoutez un ou plusieurs appels d’exécution](#create-a-mapping-of-actions-to-their-functions) de l’API [Office. actions. Associate](/javascript/api/office/office.actions#associate) pour mapper une fonction à chaque action.</span><span class="sxs-lookup"><span data-stu-id="8e830-110">[Add one or more runtime calls](#create-a-mapping-of-actions-to-their-functions) of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="8e830-111">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="8e830-111">Configure the manifest</span></span>

<span data-ttu-id="8e830-112">Il y a deux modifications mineures à effectuer par le manifeste.</span><span class="sxs-lookup"><span data-stu-id="8e830-112">There are two small changes to the manifest to make.</span></span> <span data-ttu-id="8e830-113">La première consiste à autoriser le complément à utiliser un runtime partagé et l’autre à pointer vers un fichier au format JSON où vous avez défini les raccourcis clavier.</span><span class="sxs-lookup"><span data-stu-id="8e830-113">One is to enable the add-in to use a shared runtime and the other is to point to a JSON-formatted file where you defined the keyboard shortcuts.</span></span>

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="8e830-114">Configurer le complément pour utiliser un runtime partagé</span><span class="sxs-lookup"><span data-stu-id="8e830-114">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="8e830-115">L’ajout de raccourcis clavier personnalisés nécessite que votre complément utilise le runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="8e830-115">Adding custom keyboard shortcuts requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="8e830-116">Pour plus d’informations, [configurez un complément de sorte qu’il utilise un runtime partagé](../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="8e830-116">For more information, [Configure an add-in to use a shared runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

### <a name="link-the-mapping-file-to-the-manifest"></a><span data-ttu-id="8e830-117">Lier le fichier de mappage au manifeste</span><span class="sxs-lookup"><span data-stu-id="8e830-117">Link the mapping file to the manifest</span></span>

<span data-ttu-id="8e830-118">Immédiatement en *dessous* (pas à l’intérieur) `<VersionOverrides>` de l’élément dans le manifeste, ajoutez un élément [ExtendedOverrides](../reference/manifest/extendedoverrides.md) .</span><span class="sxs-lookup"><span data-stu-id="8e830-118">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="8e830-119">Définissez l' `Url` attribut sur l’URL complète d’un fichier JSON dans votre projet que vous allez créer dans une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="8e830-119">Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a><span data-ttu-id="8e830-120">Créer ou modifier le fichier JSON des raccourcis</span><span class="sxs-lookup"><span data-stu-id="8e830-120">Create or edit the shortcuts JSON file</span></span>

<span data-ttu-id="8e830-121">Créez un fichier JSON dans votre projet.</span><span class="sxs-lookup"><span data-stu-id="8e830-121">Create a JSON file in your project.</span></span> <span data-ttu-id="8e830-122">Assurez-vous que le chemin d’accès du fichier correspond à l’emplacement que vous avez spécifié pour l' `Url` attribut de l’élément [ExtendedOverrides](../reference/manifest/extendedoverrides.md) .</span><span class="sxs-lookup"><span data-stu-id="8e830-122">Be sure the path of the file matches the location you specified for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="8e830-123">Ce fichier décrit les raccourcis clavier et les actions qu’ils vont appeler.</span><span class="sxs-lookup"><span data-stu-id="8e830-123">This file will describe your keyboard shortcuts, and the actions that they will invoke.</span></span>

1. <span data-ttu-id="8e830-124">Dans le fichier JSON, il existe deux tableaux.</span><span class="sxs-lookup"><span data-stu-id="8e830-124">Inside the JSON file, there are two arrays.</span></span> <span data-ttu-id="8e830-125">Le tableau actions contient les objets qui définissent les actions à appeler et le tableau de raccourcis contient les objets qui mappent les combinaisons de clés sur les actions.</span><span class="sxs-lookup"><span data-stu-id="8e830-125">The actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions.</span></span> <span data-ttu-id="8e830-126">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="8e830-126">Here is an example:</span></span>

    ```json
    {
        "actions": [
            {
                "id": "SHOWTASKPANE",
                "type": "ExecuteFunction",
                "name": "Show task pane for add-in"
            },
            {
                "id": "HIDETASKPANE",
                "type": "ExecuteFunction",
                "name": "Hide task pane for add-in"
            }
        ],
        "shortcuts": [
            {
                "action": "SHOWTASKPANE",
                "key": {
                    "default": "CTRL+SHIFT+UP"
                }
            },
            {
                "action": "HIDETASKPANE",
                "key": {
                    "default": "CTRL+SHIFT+DOWN"
                }
            }
        ]
    }
    ```

    <span data-ttu-id="8e830-127">Pour plus d’informations sur les objets JSON, voir [construction des objets action](#constructing-the-action-objects) et [création des objets Shortcut](#constructing-the-shortcut-objects).</span><span class="sxs-lookup"><span data-stu-id="8e830-127">For more information about the JSON objects, see [Constructing the action objects](#constructing-the-action-objects) and [Constructing the shortcut objects](#constructing-the-shortcut-objects).</span></span> <span data-ttu-id="8e830-128">Le schéma complet pour les raccourcis JSON se trouve [ surextended-manifest.schema.js](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="8e830-128">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span> <span data-ttu-id="8e830-129">(Remarque : le lien vers le schéma peut ne pas fonctionner au début de la période de préversion.)</span><span class="sxs-lookup"><span data-stu-id="8e830-129">(Note: The link to the schema may not be working early in the preview period.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="8e830-130">Vous pouvez utiliser le « contrôle » à la place de « CTRL » tout au long de cet article.</span><span class="sxs-lookup"><span data-stu-id="8e830-130">You can use "CONTROL" in place of "CTRL" throughout this article.</span></span>

    <span data-ttu-id="8e830-131">Dans une étape ultérieure, les actions seront elles-mêmes mappées vers les fonctions que vous écrivez.</span><span class="sxs-lookup"><span data-stu-id="8e830-131">In a later step, the actions will themselves be mapped to functions that you write.</span></span> <span data-ttu-id="8e830-132">Dans cet exemple, vous allez mapper ultérieurement SHOWTASKPANE à une fonction qui appelle la `Office.addin.showAsTaskpane` méthode et HIDETASKPANE à une fonction qui appelle la `Office.addin.hide` méthode.</span><span class="sxs-lookup"><span data-stu-id="8e830-132">In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.</span></span>

## <a name="create-a-mapping-of-actions-to-their-functions"></a><span data-ttu-id="8e830-133">Créer un mappage des actions sur leurs fonctions</span><span class="sxs-lookup"><span data-stu-id="8e830-133">Create a mapping of actions to their functions</span></span>

1. <span data-ttu-id="8e830-134">Dans votre projet, ouvrez le fichier JavaScript chargé par votre page HTML dans l' `<FunctionFile>` élément.</span><span class="sxs-lookup"><span data-stu-id="8e830-134">In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.</span></span>
1. <span data-ttu-id="8e830-135">Dans le fichier JavaScript, utilisez l’API [Office. actions. Associate](/javascript/api/office/office.actions#associate) pour mapper chaque action que vous avez spécifiée dans le fichier JSON à une fonction JavaScript.</span><span class="sxs-lookup"><span data-stu-id="8e830-135">In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function.</span></span> <span data-ttu-id="8e830-136">Ajoutez le code JavaScript suivant au fichier.</span><span class="sxs-lookup"><span data-stu-id="8e830-136">Add the following JavaScript to the file.</span></span> <span data-ttu-id="8e830-137">Notez ce qui suit dans le code :</span><span class="sxs-lookup"><span data-stu-id="8e830-137">Note the following about the code:</span></span>

    - <span data-ttu-id="8e830-138">Le premier paramètre est l’une des actions du fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="8e830-138">The first parameter is one of the actions from the JSON file.</span></span>
    - <span data-ttu-id="8e830-139">Le deuxième paramètre est la fonction qui s’exécute lorsqu’un utilisateur appuie sur la combinaison de touches mappée à l’action dans le fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="8e830-139">The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.</span></span>

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. <span data-ttu-id="8e830-140">Pour continuer l’exemple, utilisez `'SHOWTASKPANE'` comme premier paramètre.</span><span class="sxs-lookup"><span data-stu-id="8e830-140">To continue the example, use `'SHOWTASKPANE'` as the first parameter.</span></span>
1. <span data-ttu-id="8e830-141">Pour le corps de la fonction, utilisez la méthode [Office. AddIn. showTaskpane](/javascript/api/office/office.addin#showastaskpane--) pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="8e830-141">For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) method to open the add-in's task pane.</span></span> <span data-ttu-id="8e830-142">Une fois que vous avez fini, le code doit ressembler à ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="8e830-142">When you are done, the code should look like the following:</span></span>

    ```javascript
    Office.actions.associate('SHOWTASKPANE', function () {
        return Office.addin.showAsTaskpane()
            .then(function () {
                return;
            })
            .catch(function (error) {
                return error.code;
            });
    });
    ```

1. <span data-ttu-id="8e830-143">Ajoutez un deuxième appel de `Office.actions.associate` fonction pour mapper l' `HIDETASKPANE` action sur une fonction qui appelle [Office. AddIn. Hide](/javascript/api/office/office.addin#hide--).</span><span class="sxs-lookup"><span data-stu-id="8e830-143">Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span></span> <span data-ttu-id="8e830-144">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="8e830-144">The following is an example:</span></span>

    ```javascript
    Office.actions.associate('HIDETASKPANE', function () {
        return Office.addin.hide()
            .then(function () {
                return;
            })
            .catch(function (error) {
                return error.code;
            });
    });
    ```

<span data-ttu-id="8e830-145">Le suivi des étapes précédentes permet à votre complément de faire basculer la visibilité du volet des tâches en appuyant sur **Ctrl + Maj + Flèche vers le haut** et **Ctrl + Maj + Flèche vers le bas**.</span><span class="sxs-lookup"><span data-stu-id="8e830-145">Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Shift+Up arrow key** and **Ctrl+Shift+Down arrow key**.</span></span> <span data-ttu-id="8e830-146">Il s’agit du même comportement que celui présenté dans l' [exemple de complément Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span><span class="sxs-lookup"><span data-stu-id="8e830-146">This is the same behavior as shown in the [sample excel keyboard shortcuts add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>

## <a name="details-and-restrictions"></a><span data-ttu-id="8e830-147">Détails et restrictions</span><span class="sxs-lookup"><span data-stu-id="8e830-147">Details and restrictions</span></span>

### <a name="constructing-the-action-objects"></a><span data-ttu-id="8e830-148">Construction des objets action</span><span class="sxs-lookup"><span data-stu-id="8e830-148">Constructing the action objects</span></span>

<span data-ttu-id="8e830-149">Utilisez les instructions suivantes lorsque vous spécifiez les objets dans le `action` tableau de la shortcuts.jssur :</span><span class="sxs-lookup"><span data-stu-id="8e830-149">Use the following guidelines when specifying the objects in the `action` array of the shortcuts.json:</span></span>

- <span data-ttu-id="8e830-150">Les noms `id` de propriété et `name` sont obligatoires.</span><span class="sxs-lookup"><span data-stu-id="8e830-150">The property names `id` and `name` are mandatory.</span></span>
- <span data-ttu-id="8e830-151">La `id` propriété sert à identifier de manière unique l’action à appeler à l’aide d’un raccourci clavier.</span><span class="sxs-lookup"><span data-stu-id="8e830-151">The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.</span></span>
- <span data-ttu-id="8e830-152">La `name` propriété doit être une chaîne conviviale décrivant l’action.</span><span class="sxs-lookup"><span data-stu-id="8e830-152">The `name` property must be a user friendly string describing the action.</span></span> <span data-ttu-id="8e830-153">Il doit s’agir d’une combinaison des caractères A-Z, a-z, 0-9 et des signes de ponctuation « - », « _ » et « + ».</span><span class="sxs-lookup"><span data-stu-id="8e830-153">It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span>
- <span data-ttu-id="8e830-154">La propriété `type` est facultative.</span><span class="sxs-lookup"><span data-stu-id="8e830-154">The `type` property is optional.</span></span> <span data-ttu-id="8e830-155">Le type actuellement uniquement `ExecuteFunction` est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="8e830-155">Currently only `ExecuteFunction` type is supported.</span></span>

<span data-ttu-id="8e830-156">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="8e830-156">The following is an example:</span></span>

```json
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        },
        {
            "id": "HIDETASKPANE",
            "type": "ExecuteFunction",
            "name": "Hide task pane for add-in"
        }
    ]
```

<span data-ttu-id="8e830-157">Le schéma complet pour les raccourcis JSON se trouve [ surextended-manifest.schema.js](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="8e830-157">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span> <span data-ttu-id="8e830-158">(Remarque : le lien vers le schéma peut ne pas fonctionner au début de la période de préversion.)</span><span class="sxs-lookup"><span data-stu-id="8e830-158">(Note: The link to the schema may not be working early in the preview period.)</span></span>

### <a name="constructing-the-shortcut-objects"></a><span data-ttu-id="8e830-159">Construction des objets Shortcut</span><span class="sxs-lookup"><span data-stu-id="8e830-159">Constructing the shortcut objects</span></span>

<span data-ttu-id="8e830-160">Utilisez les instructions suivantes lorsque vous spécifiez les objets dans le `shortcuts` tableau de la shortcuts.jssur :</span><span class="sxs-lookup"><span data-stu-id="8e830-160">Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json:</span></span>

- <span data-ttu-id="8e830-161">Les noms des propriétés `action` , `key` et `default` sont obligatoires.</span><span class="sxs-lookup"><span data-stu-id="8e830-161">The property names `action`, `key`, and `default` are required.</span></span>
- <span data-ttu-id="8e830-162">La valeur de la `action` propriété est une chaîne qui doit correspondre à l’une des `id` Propriétés de l’objet action.</span><span class="sxs-lookup"><span data-stu-id="8e830-162">The value of the `action` property is a string and must match one of the `id` properties in the action object.</span></span>
- <span data-ttu-id="8e830-163">La `default` propriété peut être n’importe quelle combinaison des caractères a-z, a-z, 0-9 et des signes de ponctuation « - », « _ » et « + ».</span><span class="sxs-lookup"><span data-stu-id="8e830-163">The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span> <span data-ttu-id="8e830-164">(Par Convention, les lettres minuscules ne sont pas utilisées dans ces propriétés.)</span><span class="sxs-lookup"><span data-stu-id="8e830-164">(By convention, lower case letters are not used in these properties.)</span></span>
- <span data-ttu-id="8e830-165">La `default` propriété doit contenir le nom d’au moins une touche de modification (Alt, CTRL, Maj) et une seule autre clé.</span><span class="sxs-lookup"><span data-stu-id="8e830-165">The `default` property must contain the name of at least one modifier key (ALT, CTRL, SHIFT) and only one other key.</span></span>
- <span data-ttu-id="8e830-166">Pour Mac, nous pouvons également prendre en charge la touche de modification de commande.</span><span class="sxs-lookup"><span data-stu-id="8e830-166">For Macs, we also support the COMMAND modifier key.</span></span>
- <span data-ttu-id="8e830-167">Pour Mac, la touche ALT est mappée sur la touche d’OPTION.</span><span class="sxs-lookup"><span data-stu-id="8e830-167">For Macs, ALT is mapped to the OPTION key.</span></span> <span data-ttu-id="8e830-168">Pour Windows, la commande est mappée sur la touche CTRL.</span><span class="sxs-lookup"><span data-stu-id="8e830-168">For Windows, COMMAND is mapped to the CTRL key.</span></span>
- <span data-ttu-id="8e830-169">Lorsque deux caractères sont liés à la même clé physique dans un clavier standard, alors qu’il s’agit de synonymes dans la `default` propriété ; par exemple, Alt + a et Alt + a représentent le même raccourci, donc Ctrl +-et Ctrl + \_ car « - » et « _ » sont la même clé physique.</span><span class="sxs-lookup"><span data-stu-id="8e830-169">When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, ALT+a and ALT+A are the same shortcut, so are CTRL+- and CTRL+\_ because "-" and "_" are the same physical key.</span></span>
- <span data-ttu-id="8e830-170">Le caractère « + » indique que les touches de part et d’autre de celle-ci sont enfoncées simultanément.</span><span class="sxs-lookup"><span data-stu-id="8e830-170">The "+" character indicates that the keys on either side of it are pressed simultaneously.</span></span>

<span data-ttu-id="8e830-171">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="8e830-171">The following is an example:</span></span>

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "CTRL+SHIFT+UP"
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "CTRL+SHIFT+DOWN"
            }
        }
    ]
```

<span data-ttu-id="8e830-172">Le schéma complet pour les raccourcis JSON se trouve [ surextended-manifest.schema.js](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="8e830-172">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span> <span data-ttu-id="8e830-173">(Remarque : le lien vers le schéma peut ne pas fonctionner au début de la période de préversion.)</span><span class="sxs-lookup"><span data-stu-id="8e830-173">(Note: The link to the schema may not be working early in the preview period.)</span></span>

> [!NOTE]
> <span data-ttu-id="8e830-174">Les KeyTips, également appelés raccourcis clavier séquentiels, tels que le raccourci Excel pour choisir une couleur de remplissage **Alt + h, h**, ne sont pas pris en charge dans les compléments Office.</span><span class="sxs-lookup"><span data-stu-id="8e830-174">Keytips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office add-ins.</span></span>

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a><span data-ttu-id="8e830-175">Utilisation de raccourcis lorsque le volet Office est sélectionné</span><span class="sxs-lookup"><span data-stu-id="8e830-175">Using shortcuts when the focus is in the task pane</span></span>

<span data-ttu-id="8e830-176">Actuellement, les raccourcis clavier d’un complément Office ne peuvent être appelés que lorsque le focus de l’utilisateur se trouve dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="8e830-176">Currently, the keyboard shortcuts for an Office add-in can only be invoked when the user's focus is in the worksheet.</span></span> <span data-ttu-id="8e830-177">Lorsque le focus de l’utilisateur se trouve à l’intérieur de l’interface utilisateur Office (par exemple, le volet Office), aucun des raccourcis du complément n’est ignoré.</span><span class="sxs-lookup"><span data-stu-id="8e830-177">When the user's focus is inside the Office UI (such as the task pane), none of the add-in's shortcuts are ignored.</span></span> <span data-ttu-id="8e830-178">En guise de solution de contournement, le complément peut définir des gestionnaires de clavier qui peuvent appeler certaines actions lorsque l’utilisateur se trouve à l’intérieur de l’interface utilisateur du complément.</span><span class="sxs-lookup"><span data-stu-id="8e830-178">As a workaround, the add-in can define keyboard handlers that can invoke certain actions when the user's focus is inside of the add-in UI.</span></span>

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a><span data-ttu-id="8e830-179">Utilisation de combinaisons de touches déjà utilisées par Office ou un autre complément</span><span class="sxs-lookup"><span data-stu-id="8e830-179">Using key combinations that are already used by Office or another add-in</span></span>

<span data-ttu-id="8e830-180">Pendant la période de préversion, il n’existe aucun système permettant de déterminer ce qui se produit lorsqu’un utilisateur appuie sur une combinaison de touches enregistrée par un complément et par Office ou par un autre complément.</span><span class="sxs-lookup"><span data-stu-id="8e830-180">During the preview period, there is no system for determining what happens when a user presses a key combination that is registered by an add-in and also by Office or by another add-in.</span></span> <span data-ttu-id="8e830-181">Le comportement n’est pas défini.</span><span class="sxs-lookup"><span data-stu-id="8e830-181">Behavior is undefined.</span></span>

<span data-ttu-id="8e830-182">Actuellement, il n’existe aucune solution de contournement lorsque deux ou plusieurs compléments ont inscrit le même raccourci clavier, mais vous pouvez réduire les conflits avec Excel avec ces bonnes pratiques :</span><span class="sxs-lookup"><span data-stu-id="8e830-182">Currently, there is no workaround when two or more add-ins have registered the same keyboard shortcut, but you can minimize conflicts with Excel with these good practices:</span></span>

- <span data-ttu-id="8e830-183">Utilisez uniquement des raccourcis clavier avec le modèle suivant dans votre complément : \**Ctrl + Maj + Alt +* x \* \* \*, où *x* est une autre touche.</span><span class="sxs-lookup"><span data-stu-id="8e830-183">Use only keyboard shortcuts with the following pattern in your add-in: \**Ctrl+Shift+Alt+* x\*\*\*, where *x* is some other key.</span></span>
- <span data-ttu-id="8e830-184">Si vous avez besoin de raccourcis clavier supplémentaires, consultez la [liste des raccourcis clavier Excel](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)et évitez de les utiliser dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="8e830-184">If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.</span></span>

## <a name="browser-shortcuts-that-cannot-be-overridden"></a><span data-ttu-id="8e830-185">Raccourcis de navigateur qui ne peuvent pas être remplacés</span><span class="sxs-lookup"><span data-stu-id="8e830-185">Browser shortcuts that cannot be overridden</span></span>

<span data-ttu-id="8e830-186">Vous ne pouvez pas utiliser l’une des combinaisons de touches suivantes.</span><span class="sxs-lookup"><span data-stu-id="8e830-186">You cannot use any of the following keyboard combinations.</span></span> <span data-ttu-id="8e830-187">Elles sont utilisées par les navigateurs et ne peuvent pas être remplacées.</span><span class="sxs-lookup"><span data-stu-id="8e830-187">They are used by browsers and cannot be overridden.</span></span> <span data-ttu-id="8e830-188">Cette liste est un travail en cours.</span><span class="sxs-lookup"><span data-stu-id="8e830-188">This list is a work in progress.</span></span> <span data-ttu-id="8e830-189">Si vous découvrez d’autres combinaisons qui ne peuvent pas être remplacées, faites-le nous savoir en utilisant l’outil de commentaires en bas de cette page.</span><span class="sxs-lookup"><span data-stu-id="8e830-189">If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.</span></span>

- <span data-ttu-id="8e830-190">Ctrl + N</span><span class="sxs-lookup"><span data-stu-id="8e830-190">Ctrl+N</span></span>
- <span data-ttu-id="8e830-191">Ctrl + Maj + N</span><span class="sxs-lookup"><span data-stu-id="8e830-191">Ctrl+Shift+N</span></span>
- <span data-ttu-id="8e830-192">Ctrl + T</span><span class="sxs-lookup"><span data-stu-id="8e830-192">Ctrl+T</span></span>
- <span data-ttu-id="8e830-193">Ctrl + Maj + T</span><span class="sxs-lookup"><span data-stu-id="8e830-193">Ctrl+Shift+T</span></span>
- <span data-ttu-id="8e830-194">CTRL + W</span><span class="sxs-lookup"><span data-stu-id="8e830-194">Ctrl+W</span></span>
- <span data-ttu-id="8e830-195">Ctrl + PG. préc/PG. suiv</span><span class="sxs-lookup"><span data-stu-id="8e830-195">Ctrl+PgUp/PgDn</span></span>

## <a name="next-steps"></a><span data-ttu-id="8e830-196">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="8e830-196">Next Steps</span></span>

- <span data-ttu-id="8e830-197">Voir l’exemple de complément [Excel-clavier-raccourcis](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span><span class="sxs-lookup"><span data-stu-id="8e830-197">See the sample add-in [excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>
