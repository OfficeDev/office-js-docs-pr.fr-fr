---
title: Raccourcis clavier personnalisés dans Office des modules
description: Découvrez comment ajouter des raccourcis clavier personnalisés, également appelés combinaisons de touches, à votre Office de clavier.
ms.date: 05/05/2021
localization_priority: Normal
ms.openlocfilehash: 42c0b5190d0fc71f137284950bcb983f16845fca
ms.sourcegitcommit: 132f5082f5bf9500dad0a2eaf89d924c823e575d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/07/2021
ms.locfileid: "52266109"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins"></a><span data-ttu-id="31291-103">Ajouter des raccourcis clavier personnalisés à vos Office de travail</span><span class="sxs-lookup"><span data-stu-id="31291-103">Add custom keyboard shortcuts to your Office Add-ins</span></span>

<span data-ttu-id="31291-104">Les raccourcis clavier, également appelés combinaisons de touches, permettent aux utilisateurs de votre module de travailler plus efficacement.</span><span class="sxs-lookup"><span data-stu-id="31291-104">Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently.</span></span> <span data-ttu-id="31291-105">Les raccourcis clavier améliorent également l’accessibilité du module pour les utilisateurs présentant un handicap en offrant une alternative à la souris.</span><span class="sxs-lookup"><span data-stu-id="31291-105">Keyboard shortcuts also improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.</span></span>

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> <span data-ttu-id="31291-106">Pour commencer avec une version de travail d’un add-in avec des raccourcis clavier déjà activés, clonez et exécutez l’exemple [Excel raccourcis clavier.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)</span><span class="sxs-lookup"><span data-stu-id="31291-106">To start with a working version of an add-in with keyboard shortcuts already enabled, clone and run the sample [Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span> <span data-ttu-id="31291-107">Lorsque vous êtes prêt à ajouter des raccourcis clavier à votre propre add-in, poursuivez avec cet article.</span><span class="sxs-lookup"><span data-stu-id="31291-107">When you are ready to add keyboard shortcuts to your own add-in, continue with this article.</span></span>

<span data-ttu-id="31291-108">Il existe trois étapes pour ajouter des raccourcis clavier à un add-in :</span><span class="sxs-lookup"><span data-stu-id="31291-108">There are three steps to add keyboard shortcuts to an add-in:</span></span>

1. <span data-ttu-id="31291-109">[Configurez le manifeste du add-in.](#configure-the-manifest)</span><span class="sxs-lookup"><span data-stu-id="31291-109">[Configure the add-in's manifest](#configure-the-manifest).</span></span>
1. <span data-ttu-id="31291-110">[Créez ou modifiez le fichier JSON](#create-or-edit-the-shortcuts-json-file) de raccourcis pour définir des actions et leurs raccourcis clavier.</span><span class="sxs-lookup"><span data-stu-id="31291-110">[Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file) to define actions and their keyboard shortcuts.</span></span>
1. <span data-ttu-id="31291-111">[Ajoutez un ou plusieurs appels runtime](#create-a-mapping-of-actions-to-their-functions) de [l’API Office.actions.associate](/javascript/api/office/office.actions#associate) pour ma cartographier une fonction à chaque action.</span><span class="sxs-lookup"><span data-stu-id="31291-111">[Add one or more runtime calls](#create-a-mapping-of-actions-to-their-functions) of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="31291-112">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="31291-112">Configure the manifest</span></span>

<span data-ttu-id="31291-113">Deux petites modifications sont à apporter au manifeste.</span><span class="sxs-lookup"><span data-stu-id="31291-113">There are two small changes to the manifest to make.</span></span> <span data-ttu-id="31291-114">L’une consiste à permettre au add-in d’utiliser un runtime partagé et l’autre à pointer vers un fichier au format JSON où vous avez défini les raccourcis clavier.</span><span class="sxs-lookup"><span data-stu-id="31291-114">One is to enable the add-in to use a shared runtime and the other is to point to a JSON-formatted file where you defined the keyboard shortcuts.</span></span>

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="31291-115">Configurer le add-in pour utiliser un runtime partagé</span><span class="sxs-lookup"><span data-stu-id="31291-115">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="31291-116">L’ajout de raccourcis clavier personnalisés nécessite que votre add-in utilise le runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="31291-116">Adding custom keyboard shortcuts requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="31291-117">Pour plus d’informations, [configurez un module complémentaire pour utiliser un runtime partagé.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="31291-117">For more information, [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

### <a name="link-the-mapping-file-to-the-manifest"></a><span data-ttu-id="31291-118">Lier le fichier de mappage au manifeste</span><span class="sxs-lookup"><span data-stu-id="31291-118">Link the mapping file to the manifest</span></span>

<span data-ttu-id="31291-119">Juste *en dessous* (pas à l’intérieur) de l’élément dans le `<VersionOverrides>` manifeste, ajoutez un élément [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="31291-119">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="31291-120">Définissez l’attribut sur l’URL complète d’un fichier JSON dans votre projet que vous `Url` créerez à une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="31291-120">Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a><span data-ttu-id="31291-121">Créer ou modifier le fichier JSON de raccourcis</span><span class="sxs-lookup"><span data-stu-id="31291-121">Create or edit the shortcuts JSON file</span></span>

<span data-ttu-id="31291-122">Créez un fichier JSON dans votre projet.</span><span class="sxs-lookup"><span data-stu-id="31291-122">Create a JSON file in your project.</span></span> <span data-ttu-id="31291-123">Assurez-vous que le chemin d’accès au fichier correspond à l’emplacement que vous avez spécifié pour l’attribut de l’élément `Url` [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="31291-123">Be sure the path of the file matches the location you specified for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="31291-124">Ce fichier décrit vos raccourcis clavier et les actions qu’ils appelleront.</span><span class="sxs-lookup"><span data-stu-id="31291-124">This file will describe your keyboard shortcuts, and the actions that they will invoke.</span></span>

1. <span data-ttu-id="31291-125">Le fichier JSON se trouve à l’intérieur de deux tableaux.</span><span class="sxs-lookup"><span data-stu-id="31291-125">Inside the JSON file, there are two arrays.</span></span> <span data-ttu-id="31291-126">Le tableau d’actions contient des objets qui définissent les actions à appeler et le tableau de raccourcis contient des objets qui maient des combinaisons de touches sur des actions.</span><span class="sxs-lookup"><span data-stu-id="31291-126">The actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions.</span></span> <span data-ttu-id="31291-127">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="31291-127">Here is an example:</span></span>

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
                    "default": "Ctrl+Alt+Up"
                }
            },
            {
                "action": "HIDETASKPANE",
                "key": {
                    "default": "Ctrl+Alt+Down"
                }
            }
        ]
    }
    ```

    <span data-ttu-id="31291-128">Pour plus d’informations sur les objets JSON, voir [Construct the action objects](#construct-the-action-objects) and [Construct the shortcut objects](#construct-the-shortcut-objects).</span><span class="sxs-lookup"><span data-stu-id="31291-128">For more information about the JSON objects, see [Construct the action objects](#construct-the-action-objects) and [Construct the shortcut objects](#construct-the-shortcut-objects).</span></span> <span data-ttu-id="31291-129">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="31291-129">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

    > [!NOTE]
    > <span data-ttu-id="31291-130">Vous pouvez utiliser « CONTROL » à la place de « Ctrl » tout au long de cet article.</span><span class="sxs-lookup"><span data-stu-id="31291-130">You can use "CONTROL" in place of "Ctrl" throughout this article.</span></span>

    <span data-ttu-id="31291-131">Dans une étape ultérieure, les actions seront elles-mêmes mappées aux fonctions que vous écrivez.</span><span class="sxs-lookup"><span data-stu-id="31291-131">In a later step, the actions will themselves be mapped to functions that you write.</span></span> <span data-ttu-id="31291-132">Dans cet exemple, vous masquez ultérieurement SHOWTASKPANE à une fonction qui appelle la méthode et HIDETASKPANE à une fonction qui `Office.addin.showAsTaskpane` appelle la `Office.addin.hide` méthode.</span><span class="sxs-lookup"><span data-stu-id="31291-132">In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.</span></span>

## <a name="create-a-mapping-of-actions-to-their-functions"></a><span data-ttu-id="31291-133">Créer un mappage des actions à leurs fonctions</span><span class="sxs-lookup"><span data-stu-id="31291-133">Create a mapping of actions to their functions</span></span>

1. <span data-ttu-id="31291-134">Dans votre projet, ouvrez le fichier JavaScript chargé par votre page HTML dans `<FunctionFile>` l’élément.</span><span class="sxs-lookup"><span data-stu-id="31291-134">In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.</span></span>
1. <span data-ttu-id="31291-135">Dans le fichier JavaScript, utilisez l’API [Office.actions.associate](/javascript/api/office/office.actions#associate) pour ma cartographier chaque action que vous avez spécifiée dans le fichier JSON sur une fonction JavaScript.</span><span class="sxs-lookup"><span data-stu-id="31291-135">In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function.</span></span> <span data-ttu-id="31291-136">Ajoutez le javaScript suivant au fichier.</span><span class="sxs-lookup"><span data-stu-id="31291-136">Add the following JavaScript to the file.</span></span> <span data-ttu-id="31291-137">Notez les informations suivantes sur le code :</span><span class="sxs-lookup"><span data-stu-id="31291-137">Note the following about the code:</span></span>

    - <span data-ttu-id="31291-138">Le premier paramètre est l’une des actions du fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="31291-138">The first parameter is one of the actions from the JSON file.</span></span>
    - <span data-ttu-id="31291-139">Le deuxième paramètre est la fonction qui s’exécute lorsqu’un utilisateur appuie sur la combinaison de touches mappée à l’action dans le fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="31291-139">The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.</span></span>

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. <span data-ttu-id="31291-140">Pour continuer l’exemple, `'SHOWTASKPANE'` utilisez-le comme premier paramètre.</span><span class="sxs-lookup"><span data-stu-id="31291-140">To continue the example, use `'SHOWTASKPANE'` as the first parameter.</span></span>
1. <span data-ttu-id="31291-141">Pour le corps de la fonction, utilisez la [méthode Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) pour ouvrir le volet Des tâches du module.</span><span class="sxs-lookup"><span data-stu-id="31291-141">For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) method to open the add-in's task pane.</span></span> <span data-ttu-id="31291-142">Lorsque vous avez terminé, le code doit ressembler à ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="31291-142">When you are done, the code should look like the following:</span></span>

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

1. <span data-ttu-id="31291-143">Ajoutez un deuxième appel de fonction pour maque l’action à une `Office.actions.associate` fonction qui appelle `HIDETASKPANE` [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span><span class="sxs-lookup"><span data-stu-id="31291-143">Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span></span> <span data-ttu-id="31291-144">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="31291-144">The following is an example:</span></span>

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

<span data-ttu-id="31291-145">La suite des étapes précédentes permet à votre add-in de faire tourner la visibilité du volet Des tâches en appuyant sur **Ctrl+Alt+Haut** et **Ctrl+Alt+Bas.**</span><span class="sxs-lookup"><span data-stu-id="31291-145">Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Alt+Up** and **Ctrl+Alt+Down**.</span></span> <span data-ttu-id="31291-146">Le même comportement est illustré dans [l’exemple de raccourcis](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) clavier Excel dans le Office PnP des Office dans GitHub.</span><span class="sxs-lookup"><span data-stu-id="31291-146">The same behavior is shown in the [Excel keyboard shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) sample in the Office Add-ins PnP repo in GitHub.</span></span>

## <a name="details-and-restrictions"></a><span data-ttu-id="31291-147">Détails et restrictions</span><span class="sxs-lookup"><span data-stu-id="31291-147">Details and restrictions</span></span>

### <a name="construct-the-action-objects"></a><span data-ttu-id="31291-148">Construire les objets d’action</span><span class="sxs-lookup"><span data-stu-id="31291-148">Construct the action objects</span></span>

<span data-ttu-id="31291-149">Utilisez les instructions suivantes lors de la spécification des objets dans le tableau de la `actions` shortcuts.jssuivantes :</span><span class="sxs-lookup"><span data-stu-id="31291-149">Use the following guidelines when specifying the objects in the `actions` array of the shortcuts.json:</span></span>

- <span data-ttu-id="31291-150">Les noms des `id` propriétés `name` sont obligatoires.</span><span class="sxs-lookup"><span data-stu-id="31291-150">The property names `id` and `name` are mandatory.</span></span>
- <span data-ttu-id="31291-151">La `id` propriété est utilisée pour identifier de manière unique l’action à appeler à l’aide d’un raccourci clavier.</span><span class="sxs-lookup"><span data-stu-id="31291-151">The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.</span></span>
- <span data-ttu-id="31291-152">La `name` propriété doit être une chaîne conviviale décrivant l’action.</span><span class="sxs-lookup"><span data-stu-id="31291-152">The `name` property must be a user friendly string describing the action.</span></span> <span data-ttu-id="31291-153">Il doit s’agit d’une combinaison des caractères A - Z, a - z, 0 - 9, et des signes de ponctuation « - », « _ » et « + ».</span><span class="sxs-lookup"><span data-stu-id="31291-153">It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span>
- <span data-ttu-id="31291-154">La propriété `type` est facultative.</span><span class="sxs-lookup"><span data-stu-id="31291-154">The `type` property is optional.</span></span> <span data-ttu-id="31291-155">Actuellement, `ExecuteFunction` seul le type est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="31291-155">Currently only `ExecuteFunction` type is supported.</span></span>

<span data-ttu-id="31291-156">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="31291-156">The following is an example:</span></span>

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

<span data-ttu-id="31291-157">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="31291-157">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

### <a name="construct-the-shortcut-objects"></a><span data-ttu-id="31291-158">Construire les objets de raccourci</span><span class="sxs-lookup"><span data-stu-id="31291-158">Construct the shortcut objects</span></span>

<span data-ttu-id="31291-159">Utilisez les instructions suivantes lors de la spécification des objets dans le tableau de la `shortcuts` shortcuts.jssuivantes :</span><span class="sxs-lookup"><span data-stu-id="31291-159">Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json:</span></span>

- <span data-ttu-id="31291-160">Les noms des `action` propriétés `key` et sont `default` obligatoires.</span><span class="sxs-lookup"><span data-stu-id="31291-160">The property names `action`, `key`, and `default` are required.</span></span>
- <span data-ttu-id="31291-161">La valeur de la propriété est une chaîne et doit correspondre à l’une `action` des `id` propriétés de l’objet action.</span><span class="sxs-lookup"><span data-stu-id="31291-161">The value of the `action` property is a string and must match one of the `id` properties in the action object.</span></span>
- <span data-ttu-id="31291-162">La propriété peut être n’importe quelle combinaison des caractères `default` A - Z, -z, 0 - 9 et les signes de ponctuation « - », « _ » et « + ».</span><span class="sxs-lookup"><span data-stu-id="31291-162">The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span> <span data-ttu-id="31291-163">(Par convention, les lettres majuscules ne sont pas utilisées dans ces propriétés.)</span><span class="sxs-lookup"><span data-stu-id="31291-163">(By convention, lower case letters are not used in these properties.)</span></span>
- <span data-ttu-id="31291-164">La propriété doit contenir le nom d’au moins une touche de `default` modification (Alt, Ctrl, Shift) et une seule autre touche.</span><span class="sxs-lookup"><span data-stu-id="31291-164">The `default` property must contain the name of at least one modifier key (Alt, Ctrl, Shift) and only one other key.</span></span>
- <span data-ttu-id="31291-165">Pour les Mac, nous pris en charge également la touche Modificateur de commande.</span><span class="sxs-lookup"><span data-stu-id="31291-165">For Macs, we also support the Command modifier key.</span></span>
- <span data-ttu-id="31291-166">Pour les Mac, Alt est mappé à la touche Option.</span><span class="sxs-lookup"><span data-stu-id="31291-166">For Macs, Alt is mapped to the Option key.</span></span> <span data-ttu-id="31291-167">Pour Windows, La commande est mappée sur la touche Ctrl.</span><span class="sxs-lookup"><span data-stu-id="31291-167">For Windows, Command is mapped to the Ctrl key.</span></span>
- <span data-ttu-id="31291-168">Lorsque deux caractères sont liés à la même touche physique dans un clavier standard, ils sont synonymes dans la propriété ; par exemple, Alt+a et Alt+A sont le même raccourci, tout comme `default` Ctrl+- et Ctrl+ car « - » et « _ » sont la même touche \_ physique.</span><span class="sxs-lookup"><span data-stu-id="31291-168">When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, Alt+a and Alt+A are the same shortcut, so are Ctrl+- and Ctrl+\_ because "-" and "_" are the same physical key.</span></span>
- <span data-ttu-id="31291-169">Le caractère « + » indique que les touches de chaque côté de celui-ci sont entrées simultanément.</span><span class="sxs-lookup"><span data-stu-id="31291-169">The "+" character indicates that the keys on either side of it are pressed simultaneously.</span></span>

<span data-ttu-id="31291-170">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="31291-170">The following is an example:</span></span>

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "Ctrl+Alt+Up"
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "Ctrl+Alt+Down"
            }
        }
    ]
```

<span data-ttu-id="31291-171">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="31291-171">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!NOTE]
> <span data-ttu-id="31291-172">Les touches d’accès, également appelées raccourcis de touches séquentiels, tels que le raccourci Excel pour choisir une couleur de remplissage **Alt+H, H,** ne sont pas pris en charge dans les Office.</span><span class="sxs-lookup"><span data-stu-id="31291-172">KeyTips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office Add-ins.</span></span>

## <a name="avoid-key-combinations-in-use-by-other-add-ins"></a><span data-ttu-id="31291-173">Éviter les combinaisons de touches en cours d’utilisation par d’autres modules</span><span class="sxs-lookup"><span data-stu-id="31291-173">Avoid key combinations in use by other add-ins</span></span>

<span data-ttu-id="31291-174">De nombreux raccourcis clavier sont déjà utilisés par les Office.</span><span class="sxs-lookup"><span data-stu-id="31291-174">There are many keyboard shortcuts that are already in use by Office.</span></span> <span data-ttu-id="31291-175">Évitez d’inscrire des raccourcis clavier pour votre module qui sont déjà utilisés. Cependant, dans certains cas, il peut être nécessaire de remplacer les raccourcis clavier existants ou de gérer les conflits entre plusieurs modules qui ont inscrit le même raccourci clavier.</span><span class="sxs-lookup"><span data-stu-id="31291-175">Avoid registering keyboard shortcuts for your add-in that are already in use, however there may be some instances where it is necessary to override existing keyboard shortcuts or handle conflicts between multiple add-ins that have registered the same keyboard shortcut.</span></span>

<span data-ttu-id="31291-176">En cas de conflit, l’utilisateur voit une boîte de dialogue la première fois qu’il tente d’utiliser un raccourci clavier en conflit, notez que le nom de l’action qui s’affiche dans cette boîte de dialogue est la propriété de l’objet action dans le `name` `shortcuts.json` fichier.</span><span class="sxs-lookup"><span data-stu-id="31291-176">In the case of a conflict, the user will see a dialog box the first time they attempt to use a conflicting keyboard shortcut, note that the action name that is displayed in this dialog is the `name` property in the action object in `shortcuts.json` file.</span></span>

![Illustration montrant un conflit modal avec deux actions différentes pour un seul raccourci](../images/add-in-shortcut-conflict-modal.png)

<span data-ttu-id="31291-178">L’utilisateur peut sélectionner l’action que le raccourci clavier va prendre.</span><span class="sxs-lookup"><span data-stu-id="31291-178">The user can select which action the keyboard shortcut will take.</span></span> <span data-ttu-id="31291-179">Après avoir fait la sélection, la préférence est enregistrée pour les futures utilisations du même raccourci.</span><span class="sxs-lookup"><span data-stu-id="31291-179">After making the selection, the preference is saved for future uses of the same shortcut.</span></span> <span data-ttu-id="31291-180">Les préférences de raccourci sont enregistrées par utilisateur, par plateforme.</span><span class="sxs-lookup"><span data-stu-id="31291-180">The shortcut preferences are saved per user, per platform.</span></span> <span data-ttu-id="31291-181">Si l’utilisateur souhaite modifier ses préférences,  il peut appeler la commande Réinitialiser les préférences de raccourci des Office dans la zone de recherche **Rechercher.**</span><span class="sxs-lookup"><span data-stu-id="31291-181">If the user wishes to change their preferences, they can invoke the **Reset Office Add-ins shortcut preferences** command from the **Tell me** search box.</span></span> <span data-ttu-id="31291-182">L’appel de la commande permet d’effacer toutes les préférences de raccourci de l’utilisateur et l’utilisateur sera de nouveau invité à utiliser la boîte de dialogue de conflit la prochaine fois qu’il tentera d’utiliser un raccourci conflictuelle :</span><span class="sxs-lookup"><span data-stu-id="31291-182">Invoking the command clears all of the user's add-in shortcut preferences and the user will again be prompted with the conflict dialog box the next time they attempt to use a conflicting shortcut:</span></span>

![Zone de recherche Rechercher dans Excel l’action de réinitialisation Office des préférences de raccourci de l’utilisateur](../images/add-in-reset-shortcuts-action.png)

<span data-ttu-id="31291-184">Pour une expérience utilisateur de qualité, nous vous recommandons de minimiser les conflits avec Excel avec ces bonnes pratiques :</span><span class="sxs-lookup"><span data-stu-id="31291-184">For the best user experience, we recommend that you minimize conflicts with Excel with these good practices:</span></span>

- <span data-ttu-id="31291-185">Utilisez uniquement les raccourcis clavier avec le modèle suivant : \**Ctrl+Shift+Alt+* x\*\*\*, où *x* est une autre touche.</span><span class="sxs-lookup"><span data-stu-id="31291-185">Use only keyboard shortcuts with the following pattern: \**Ctrl+Shift+Alt+* x\*\*\*, where *x* is some other key.</span></span>
- <span data-ttu-id="31291-186">Si vous avez besoin de raccourcis clavier, consultez la liste des [raccourcis](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)clavier Excel et évitez d’en utiliser un dans votre module.</span><span class="sxs-lookup"><span data-stu-id="31291-186">If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.</span></span>
- <span data-ttu-id="31291-187">Lorsque le focus du clavier se trouve à l’intérieur de l’interface utilisateur du module, **Ctrl+Espace et** **Ctrl+Shift+F10** ne fonctionnent pas, car il s’agit de raccourcis d’accessibilité essentiels.</span><span class="sxs-lookup"><span data-stu-id="31291-187">When the keyboard focus is inside the add-in UI, **Ctrl+Spacebar** and **Ctrl+Shift+F10** will not work as these are essential accessibility shortcuts.</span></span>
- <span data-ttu-id="31291-188">Sur un ordinateur Windows ou Mac, si la commande « Réinitialiser les préférences de raccourci des macros de Office » n’est pas disponible dans le menu de recherche, l’utilisateur peut ajouter manuellement la commande au ruban en personnalisant le ruban via le menu contexté.</span><span class="sxs-lookup"><span data-stu-id="31291-188">On a Windows or Mac computer, if the "Reset Office Add-ins shortcut preferences" command is not available on the search menu, the user can manually add the command to the ribbon by customizing the ribbon through the context menu.</span></span>

## <a name="customize-the-keyboard-shortcuts-per-platform"></a><span data-ttu-id="31291-189">Personnaliser les raccourcis clavier par plateforme</span><span class="sxs-lookup"><span data-stu-id="31291-189">Customize the keyboard shortcuts per platform</span></span>

<span data-ttu-id="31291-190">Il est possible de personnaliser les raccourcis pour qu’ils soient spécifiques à la plateforme.</span><span class="sxs-lookup"><span data-stu-id="31291-190">It's possible to customize shortcuts to be platform-specific.</span></span> <span data-ttu-id="31291-191">Voici un exemple de l’objet qui personnalise les raccourcis pour chacune des `shortcuts` plateformes suivantes : `windows` , , `mac` `web` .</span><span class="sxs-lookup"><span data-stu-id="31291-191">The following is an example of the `shortcuts` object that customizes the shortcuts for each of the following platforms: `windows`, `mac`, `web`.</span></span> <span data-ttu-id="31291-192">Notez que vous devez toujours avoir une touche `default` de raccourci pour chaque raccourci.</span><span class="sxs-lookup"><span data-stu-id="31291-192">Note that you must still have a `default` shortcut key for each shortcut.</span></span>

<span data-ttu-id="31291-193">Dans l’exemple suivant, la clé est la clé de retour pour toute `default` plateforme qui n’est pas spécifiée.</span><span class="sxs-lookup"><span data-stu-id="31291-193">In the following example, the `default` key is the fallback key for any platform that is not specified.</span></span> <span data-ttu-id="31291-194">La seule plateforme non spécifiée est Windows, donc la `default` clé s’applique uniquement aux Windows.</span><span class="sxs-lookup"><span data-stu-id="31291-194">The only platform not specified is Windows, so the `default` key will only apply to Windows.</span></span>

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "Ctrl+Alt+Up",
                "mac": "Command+Shift+Up",
                "web": "Ctrl+Alt+1",
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "Ctrl+Alt+Down",
                "mac": "Command+Shift+Down",
                "web": "Ctrl+Alt+2"
            }
        }
    ]
```

## <a name="localize-the-keyboard-shortcuts-json"></a><span data-ttu-id="31291-195">Localisez les raccourcis clavier JSON</span><span class="sxs-lookup"><span data-stu-id="31291-195">Localize the keyboard shortcuts JSON</span></span>

<span data-ttu-id="31291-196">Si votre add-in prend en charge plusieurs paramètres régionaux, vous devez trouver la propriété des `name` objets d’action.</span><span class="sxs-lookup"><span data-stu-id="31291-196">If your add-in supports multiple locales, you'll need to localize the `name` property of the action objects.</span></span> <span data-ttu-id="31291-197">En outre, si l’un des paramètres régionaux que le add-in prend en charge a des alphabets ou des systèmes d’écriture différents, et par conséquent différents claviers, vous devrez peut-être également trouver les raccourcis.</span><span class="sxs-lookup"><span data-stu-id="31291-197">Also, if any of the locales that the add-in supports have alphabets or different writing systems, and hence different keyboards, you may need to localize the shortcuts also.</span></span> <span data-ttu-id="31291-198">Pour plus d’informations sur la façon de trouver les raccourcis clavier JSON, voir [Localize extended overrides](../develop/localization.md#localize-extended-overrides).</span><span class="sxs-lookup"><span data-stu-id="31291-198">For information about how to localize the keyboard shortcuts JSON, see [Localize extended overrides](../develop/localization.md#localize-extended-overrides).</span></span>

## <a name="browser-shortcuts-that-cannot-be-overridden"></a><span data-ttu-id="31291-199">Raccourcis du navigateur qui ne peuvent pas être préférés</span><span class="sxs-lookup"><span data-stu-id="31291-199">Browser shortcuts that cannot be overridden</span></span>

<span data-ttu-id="31291-200">Lorsque vous utilisez des raccourcis clavier personnalisés sur le web, certains raccourcis clavier utilisés par le navigateur ne peuvent pas être préférés par les modules. Cette liste est un travail en cours.</span><span class="sxs-lookup"><span data-stu-id="31291-200">When using custom keyboard shortcuts on the web, some keyboard shortcuts that are used by the browser cannot be overridden by add-ins. This list is a work in progress.</span></span> <span data-ttu-id="31291-201">Si vous découvrez d’autres combinaisons qui ne peuvent pas être overridées, faites-le nous savoir à l’aide de l’outil de commentaires en bas de cette page.</span><span class="sxs-lookup"><span data-stu-id="31291-201">If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.</span></span>

- <span data-ttu-id="31291-202">Ctrl+N</span><span class="sxs-lookup"><span data-stu-id="31291-202">Ctrl+N</span></span>
- <span data-ttu-id="31291-203">Ctrl+Shift+N</span><span class="sxs-lookup"><span data-stu-id="31291-203">Ctrl+Shift+N</span></span>
- <span data-ttu-id="31291-204">Ctrl+T</span><span class="sxs-lookup"><span data-stu-id="31291-204">Ctrl+T</span></span>
- <span data-ttu-id="31291-205">Ctrl+Shift+T</span><span class="sxs-lookup"><span data-stu-id="31291-205">Ctrl+Shift+T</span></span>
- <span data-ttu-id="31291-206">Ctrl+W</span><span class="sxs-lookup"><span data-stu-id="31291-206">Ctrl+W</span></span>
- <span data-ttu-id="31291-207">Ctrl+PgUp/PgDn</span><span class="sxs-lookup"><span data-stu-id="31291-207">Ctrl+PgUp/PgDn</span></span>

## <a name="next-steps"></a><span data-ttu-id="31291-208">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="31291-208">Next Steps</span></span>

- <span data-ttu-id="31291-209">Consultez [l Excel exemple de raccourcis](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) clavier.</span><span class="sxs-lookup"><span data-stu-id="31291-209">See the [Excel keyboard shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) sample add-in.</span></span>
- <span data-ttu-id="31291-210">Obtenez une vue d’ensemble de l’utilisation des substitutions étendues dans [Work avec des substitutions étendues du manifeste.](../develop/extended-overrides.md)</span><span class="sxs-lookup"><span data-stu-id="31291-210">Get an overview of working with extended overrides in [Work with extended overrides of the manifest](../develop/extended-overrides.md).</span></span>
