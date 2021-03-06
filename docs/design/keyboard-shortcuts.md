---
title: Raccourcis clavier personnalisés dans les add-ins Office
description: Découvrez comment ajouter des raccourcis clavier personnalisés, également appelés combinaisons de touches, à votre add-in Office.
ms.date: 02/02/2021
localization_priority: Normal
ms.openlocfilehash: c767c6d5bc23f0a44422452839cd8bdf87bd8715
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505198"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a><span data-ttu-id="6a725-103">Ajouter des raccourcis clavier personnalisés à vos add-ins Office (aperçu)</span><span class="sxs-lookup"><span data-stu-id="6a725-103">Add Custom keyboard shortcuts to your Office Add-ins (preview)</span></span>

<span data-ttu-id="6a725-104">Les raccourcis clavier, également appelés combinaisons de touches, permettent aux utilisateurs de votre module de travailler plus efficacement et améliorent l’accessibilité du module pour les utilisateurs présentant un handicap en offrant une alternative à la souris.</span><span class="sxs-lookup"><span data-stu-id="6a725-104">Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently and they improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.</span></span>

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> <span data-ttu-id="6a725-105">Pour commencer avec une version de travail d’un add-in avec des raccourcis clavier déjà activés, clonez et exécutez l’exemple de [raccourcis clavier Excel.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)</span><span class="sxs-lookup"><span data-stu-id="6a725-105">To start with a working version of an add-in with keyboard shortcuts already enabled, clone and run the sample [Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span> <span data-ttu-id="6a725-106">Lorsque vous êtes prêt à ajouter des raccourcis clavier à votre propre add-in, poursuivez avec cet article.</span><span class="sxs-lookup"><span data-stu-id="6a725-106">When you are ready to add keyboard shortcuts to your own add-in, continue with this article.</span></span>

<span data-ttu-id="6a725-107">Il existe trois étapes pour ajouter des raccourcis clavier à un add-in :</span><span class="sxs-lookup"><span data-stu-id="6a725-107">There are three steps to add keyboard shortcuts to an add-in:</span></span>

1. <span data-ttu-id="6a725-108">[Configurez le manifeste du add-in.](#configure-the-manifest)</span><span class="sxs-lookup"><span data-stu-id="6a725-108">[Configure the add-in's manifest](#configure-the-manifest).</span></span>
1. <span data-ttu-id="6a725-109">[Créez ou modifiez le fichier JSON](#create-or-edit-the-shortcuts-json-file) de raccourcis pour définir des actions et leurs raccourcis clavier.</span><span class="sxs-lookup"><span data-stu-id="6a725-109">[Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file) to define actions and their keyboard shortcuts.</span></span>
1. <span data-ttu-id="6a725-110">[Ajoutez un ou plusieurs appels d’runtime](#create-a-mapping-of-actions-to-their-functions) de l’API [Office.actions.associate](/javascript/api/office/office.actions#associate) pour ma cartographier une fonction à chaque action.</span><span class="sxs-lookup"><span data-stu-id="6a725-110">[Add one or more runtime calls](#create-a-mapping-of-actions-to-their-functions) of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="6a725-111">Configurer le manifeste</span><span class="sxs-lookup"><span data-stu-id="6a725-111">Configure the manifest</span></span>

<span data-ttu-id="6a725-112">Deux petites modifications sont à apporter au manifeste.</span><span class="sxs-lookup"><span data-stu-id="6a725-112">There are two small changes to the manifest to make.</span></span> <span data-ttu-id="6a725-113">L’une consiste à permettre au add-in d’utiliser un runtime partagé et l’autre à pointer vers un fichier au format JSON où vous avez défini les raccourcis clavier.</span><span class="sxs-lookup"><span data-stu-id="6a725-113">One is to enable the add-in to use a shared runtime and the other is to point to a JSON-formatted file where you defined the keyboard shortcuts.</span></span>

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="6a725-114">Configurer le add-in pour utiliser un runtime partagé</span><span class="sxs-lookup"><span data-stu-id="6a725-114">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="6a725-115">L’ajout de raccourcis clavier personnalisés nécessite que votre add-in utilise le runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="6a725-115">Adding custom keyboard shortcuts requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="6a725-116">Pour plus d’informations, [configurez un module complémentaire pour utiliser un runtime partagé.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="6a725-116">For more information, [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

### <a name="link-the-mapping-file-to-the-manifest"></a><span data-ttu-id="6a725-117">Lier le fichier de mappage au manifeste</span><span class="sxs-lookup"><span data-stu-id="6a725-117">Link the mapping file to the manifest</span></span>

<span data-ttu-id="6a725-118">Juste *en dessous* (pas à l’intérieur) de l’élément dans le manifeste, ajoutez un élément `<VersionOverrides>` [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="6a725-118">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="6a725-119">Définissez l’attribut sur l’URL complète d’un fichier JSON dans votre projet que vous `Url` créerez à une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="6a725-119">Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a><span data-ttu-id="6a725-120">Créer ou modifier le fichier JSON de raccourcis</span><span class="sxs-lookup"><span data-stu-id="6a725-120">Create or edit the shortcuts JSON file</span></span>

<span data-ttu-id="6a725-121">Créez un fichier JSON dans votre projet.</span><span class="sxs-lookup"><span data-stu-id="6a725-121">Create a JSON file in your project.</span></span> <span data-ttu-id="6a725-122">Assurez-vous que le chemin d’accès au fichier correspond à l’emplacement que vous avez spécifié pour l’attribut de l’élément `Url` [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="6a725-122">Be sure the path of the file matches the location you specified for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="6a725-123">Ce fichier décrit vos raccourcis clavier et les actions qu’ils appelleront.</span><span class="sxs-lookup"><span data-stu-id="6a725-123">This file will describe your keyboard shortcuts, and the actions that they will invoke.</span></span>

1. <span data-ttu-id="6a725-124">Le fichier JSON se trouve à l’intérieur de deux tableaux.</span><span class="sxs-lookup"><span data-stu-id="6a725-124">Inside the JSON file, there are two arrays.</span></span> <span data-ttu-id="6a725-125">Le tableau d’actions contient des objets qui définissent les actions à appeler et le tableau de raccourcis contient des objets qui maient des combinaisons de touches sur des actions.</span><span class="sxs-lookup"><span data-stu-id="6a725-125">The actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions.</span></span> <span data-ttu-id="6a725-126">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="6a725-126">Here is an example:</span></span>

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

    <span data-ttu-id="6a725-127">Pour plus d’informations sur les objets JSON, voir [Constructing the action objects](#constructing-the-action-objects) and [Constructing the shortcut objects](#constructing-the-shortcut-objects).</span><span class="sxs-lookup"><span data-stu-id="6a725-127">For more information about the JSON objects, see [Constructing the action objects](#constructing-the-action-objects) and [Constructing the shortcut objects](#constructing-the-shortcut-objects).</span></span> <span data-ttu-id="6a725-128">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="6a725-128">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

    > [!NOTE]
    > <span data-ttu-id="6a725-129">Vous pouvez utiliser « CONTROL » à la place de « Ctrl » tout au long de cet article.</span><span class="sxs-lookup"><span data-stu-id="6a725-129">You can use "CONTROL" in place of "CTRL" throughout this article.</span></span>

    <span data-ttu-id="6a725-130">Dans une étape ultérieure, les actions seront elles-mêmes mappées aux fonctions que vous écrivez.</span><span class="sxs-lookup"><span data-stu-id="6a725-130">In a later step, the actions will themselves be mapped to functions that you write.</span></span> <span data-ttu-id="6a725-131">Dans cet exemple, vous masquez ultérieurement SHOWTASKPANE à une fonction qui appelle la méthode et HIDETASKPANE à une fonction qui `Office.addin.showAsTaskpane` appelle la `Office.addin.hide` méthode.</span><span class="sxs-lookup"><span data-stu-id="6a725-131">In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.</span></span>

## <a name="create-a-mapping-of-actions-to-their-functions"></a><span data-ttu-id="6a725-132">Créer un mappage des actions à leurs fonctions</span><span class="sxs-lookup"><span data-stu-id="6a725-132">Create a mapping of actions to their functions</span></span>

1. <span data-ttu-id="6a725-133">Dans votre projet, ouvrez le fichier JavaScript chargé par votre page HTML dans `<FunctionFile>` l’élément.</span><span class="sxs-lookup"><span data-stu-id="6a725-133">In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.</span></span>
1. <span data-ttu-id="6a725-134">Dans le fichier JavaScript, utilisez l’API [Office.actions.associate](/javascript/api/office/office.actions#associate) pour ma cartographier chaque action que vous avez spécifiée dans le fichier JSON sur une fonction JavaScript.</span><span class="sxs-lookup"><span data-stu-id="6a725-134">In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function.</span></span> <span data-ttu-id="6a725-135">Ajoutez le javaScript suivant au fichier.</span><span class="sxs-lookup"><span data-stu-id="6a725-135">Add the following JavaScript to the file.</span></span> <span data-ttu-id="6a725-136">Notez ce qui suit à propos du code :</span><span class="sxs-lookup"><span data-stu-id="6a725-136">Note the following about the code:</span></span>

    - <span data-ttu-id="6a725-137">Le premier paramètre est l’une des actions du fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="6a725-137">The first parameter is one of the actions from the JSON file.</span></span>
    - <span data-ttu-id="6a725-138">Le deuxième paramètre est la fonction qui s’exécute lorsqu’un utilisateur appuie sur la combinaison de touches mappée à l’action dans le fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="6a725-138">The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.</span></span>

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. <span data-ttu-id="6a725-139">Pour continuer l’exemple, `'SHOWTASKPANE'` utilisez-le comme premier paramètre.</span><span class="sxs-lookup"><span data-stu-id="6a725-139">To continue the example, use `'SHOWTASKPANE'` as the first parameter.</span></span>
1. <span data-ttu-id="6a725-140">Pour le corps de la fonction, utilisez la méthode [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) pour ouvrir le volet Office du add-in.</span><span class="sxs-lookup"><span data-stu-id="6a725-140">For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) method to open the add-in's task pane.</span></span> <span data-ttu-id="6a725-141">Lorsque vous avez terminé, le code doit ressembler à ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="6a725-141">When you are done, the code should look like the following:</span></span>

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

1. <span data-ttu-id="6a725-142">Ajoutez un deuxième appel de `Office.actions.associate` fonction pour maque `HIDETASKPANE` l’action sur une fonction qui appelle [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span><span class="sxs-lookup"><span data-stu-id="6a725-142">Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span></span> <span data-ttu-id="6a725-143">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="6a725-143">The following is an example:</span></span>

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

<span data-ttu-id="6a725-144">La suite des étapes précédentes permet à votre add-in de faire tourner la visibilité du volet Des tâches en appuyant sur **Ctrl+Shift+Flèche** vers le haut et **Ctrl+Shift+Flèche** vers le bas.</span><span class="sxs-lookup"><span data-stu-id="6a725-144">Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Shift+Up arrow key** and **Ctrl+Shift+Down arrow key**.</span></span> <span data-ttu-id="6a725-145">Il s’agit du même comportement que dans l’exemple de [raccourcis clavier Excel.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)</span><span class="sxs-lookup"><span data-stu-id="6a725-145">This is the same behavior as shown in the [sample excel keyboard shortcuts add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>

## <a name="details-and-restrictions"></a><span data-ttu-id="6a725-146">Détails et restrictions</span><span class="sxs-lookup"><span data-stu-id="6a725-146">Details and restrictions</span></span>

### <a name="constructing-the-action-objects"></a><span data-ttu-id="6a725-147">Construction des objets d’action</span><span class="sxs-lookup"><span data-stu-id="6a725-147">Constructing the action objects</span></span>

<span data-ttu-id="6a725-148">Utilisez les instructions suivantes lors de la spécification des objets dans le tableau de la `action` shortcuts.jssuivantes :</span><span class="sxs-lookup"><span data-stu-id="6a725-148">Use the following guidelines when specifying the objects in the `action` array of the shortcuts.json:</span></span>

- <span data-ttu-id="6a725-149">Les noms des `id` propriétés `name` sont obligatoires.</span><span class="sxs-lookup"><span data-stu-id="6a725-149">The property names `id` and `name` are mandatory.</span></span>
- <span data-ttu-id="6a725-150">La `id` propriété est utilisée pour identifier de manière unique l’action à appeler à l’aide d’un raccourci clavier.</span><span class="sxs-lookup"><span data-stu-id="6a725-150">The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.</span></span>
- <span data-ttu-id="6a725-151">La `name` propriété doit être une chaîne conviviale décrivant l’action.</span><span class="sxs-lookup"><span data-stu-id="6a725-151">The `name` property must be a user friendly string describing the action.</span></span> <span data-ttu-id="6a725-152">Il doit s’agit d’une combinaison des caractères A - Z, a - z, 0 - 9 et des signes de ponctuation « - », « _ » et « + ».</span><span class="sxs-lookup"><span data-stu-id="6a725-152">It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span>
- <span data-ttu-id="6a725-153">La propriété `type` est facultative.</span><span class="sxs-lookup"><span data-stu-id="6a725-153">The `type` property is optional.</span></span> <span data-ttu-id="6a725-154">Actuellement, `ExecuteFunction` seul le type est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="6a725-154">Currently only `ExecuteFunction` type is supported.</span></span>

<span data-ttu-id="6a725-155">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="6a725-155">The following is an example:</span></span>

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

<span data-ttu-id="6a725-156">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="6a725-156">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

### <a name="constructing-the-shortcut-objects"></a><span data-ttu-id="6a725-157">Construction des objets de raccourci</span><span class="sxs-lookup"><span data-stu-id="6a725-157">Constructing the shortcut objects</span></span>

<span data-ttu-id="6a725-158">Utilisez les instructions suivantes lors de la spécification des objets dans le tableau de la `shortcuts` shortcuts.jssuivantes :</span><span class="sxs-lookup"><span data-stu-id="6a725-158">Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json:</span></span>

- <span data-ttu-id="6a725-159">Les noms des `action` propriétés `key` et sont `default` obligatoires.</span><span class="sxs-lookup"><span data-stu-id="6a725-159">The property names `action`, `key`, and `default` are required.</span></span>
- <span data-ttu-id="6a725-160">La valeur de la propriété est une chaîne et doit correspondre à l’une `action` des `id` propriétés de l’objet action.</span><span class="sxs-lookup"><span data-stu-id="6a725-160">The value of the `action` property is a string and must match one of the `id` properties in the action object.</span></span>
- <span data-ttu-id="6a725-161">La propriété peut être n’importe quelle combinaison des caractères `default` A - Z, -z, 0 - 9 et les signes de ponctuation « - », « _ » et « + ».</span><span class="sxs-lookup"><span data-stu-id="6a725-161">The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span> <span data-ttu-id="6a725-162">(Par convention, les lettres majuscules ne sont pas utilisées dans ces propriétés.)</span><span class="sxs-lookup"><span data-stu-id="6a725-162">(By convention, lower case letters are not used in these properties.)</span></span>
- <span data-ttu-id="6a725-163">La propriété doit contenir le nom d’au moins une touche de `default` modification (ALT, Ctrl, SHIFT) et une seule autre touche.</span><span class="sxs-lookup"><span data-stu-id="6a725-163">The `default` property must contain the name of at least one modifier key (ALT, CTRL, SHIFT) and only one other key.</span></span>
- <span data-ttu-id="6a725-164">Pour les Mac, nous prise en charge également la touche modificateur COMMAND.</span><span class="sxs-lookup"><span data-stu-id="6a725-164">For Macs, we also support the COMMAND modifier key.</span></span>
- <span data-ttu-id="6a725-165">Pour les Mac, Alt est mappée sur la touche OPTION.</span><span class="sxs-lookup"><span data-stu-id="6a725-165">For Macs, ALT is mapped to the OPTION key.</span></span> <span data-ttu-id="6a725-166">Pour Windows, la commande est mappée sur la touche Ctrl.</span><span class="sxs-lookup"><span data-stu-id="6a725-166">For Windows, COMMAND is mapped to the CTRL key.</span></span>
- <span data-ttu-id="6a725-167">Lorsque deux caractères sont liés à la même touche physique dans un clavier standard, ils sont synonymes dans la propriété ; par exemple, ALT+a et ALT+A sont le même raccourci, tout comme `default` Ctrl+- et Ctrl+ car « - » et « _ » sont la même touche \_ physique.</span><span class="sxs-lookup"><span data-stu-id="6a725-167">When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, ALT+a and ALT+A are the same shortcut, so are CTRL+- and CTRL+\_ because "-" and "_" are the same physical key.</span></span>
- <span data-ttu-id="6a725-168">Le caractère « + » indique que les touches de chaque côté de celui-ci sont entrées simultanément.</span><span class="sxs-lookup"><span data-stu-id="6a725-168">The "+" character indicates that the keys on either side of it are pressed simultaneously.</span></span>

<span data-ttu-id="6a725-169">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="6a725-169">The following is an example:</span></span>

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

<span data-ttu-id="6a725-170">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="6a725-170">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!NOTE]
> <span data-ttu-id="6a725-171">Les touches d’accès, également appelées raccourcis clavier séquentiels, tels que le raccourci Excel pour choisir une couleur de remplissage **Alt+H, H,** ne sont pas pris en charge dans les add-ins Office.</span><span class="sxs-lookup"><span data-stu-id="6a725-171">Keytips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office Add-ins.</span></span>

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a><span data-ttu-id="6a725-172">Utilisation de raccourcis lorsque le focus se trouve dans le volet Des tâches</span><span class="sxs-lookup"><span data-stu-id="6a725-172">Using shortcuts when the focus is in the task pane</span></span>

<span data-ttu-id="6a725-173">Actuellement, les raccourcis clavier d’un add-in Office ne peuvent être appelés que lorsque le focus de l’utilisateur se trouve dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="6a725-173">Currently, the keyboard shortcuts for an Office Add-in can only be invoked when the user's focus is in the worksheet.</span></span> <span data-ttu-id="6a725-174">Lorsque le focus de l’utilisateur se trouve à l’intérieur de l’interface utilisateur d’Office (par exemple, le volet Office), aucun des raccourcis du add-in n’est ignoré.</span><span class="sxs-lookup"><span data-stu-id="6a725-174">When the user's focus is inside the Office UI (such as the task pane), none of the add-in's shortcuts are ignored.</span></span> <span data-ttu-id="6a725-175">Comme solution de contournement, le add-in peut définir des handlers de clavier qui peuvent appeler certaines actions lorsque le focus de l’utilisateur est à l’intérieur de l’interface utilisateur du add-in.</span><span class="sxs-lookup"><span data-stu-id="6a725-175">As a workaround, the add-in can define keyboard handlers that can invoke certain actions when the user's focus is inside of the add-in UI.</span></span>

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a><span data-ttu-id="6a725-176">Utilisation de combinaisons de touches déjà utilisées par Office ou un autre module</span><span class="sxs-lookup"><span data-stu-id="6a725-176">Using key combinations that are already used by Office or another add-in</span></span>

<span data-ttu-id="6a725-177">Pendant la période d’aperçu, il n’existe aucun système permettant de déterminer ce qui se produit lorsqu’un utilisateur appuie sur une combinaison de touches inscrite par un module et par Office ou par un autre.</span><span class="sxs-lookup"><span data-stu-id="6a725-177">During the preview period, there is no system for determining what happens when a user presses a key combination that is registered by an add-in and also by Office or by another add-in.</span></span> <span data-ttu-id="6a725-178">Le comportement n’est pas définie.</span><span class="sxs-lookup"><span data-stu-id="6a725-178">Behavior is undefined.</span></span>

<span data-ttu-id="6a725-179">Actuellement, il n’existe aucune solution de contournement lorsque deux ou plusieurs modules ont inscrit le même raccourci clavier, mais vous pouvez réduire les conflits avec Excel avec ces bonnes pratiques :</span><span class="sxs-lookup"><span data-stu-id="6a725-179">Currently, there is no workaround when two or more add-ins have registered the same keyboard shortcut, but you can minimize conflicts with Excel with these good practices:</span></span>

- <span data-ttu-id="6a725-180">Utilisez uniquement les raccourcis clavier avec le modèle suivant dans votre add-in : \**Ctrl+Shift+Alt+* x\*\*\*, où *x* est une autre touche.</span><span class="sxs-lookup"><span data-stu-id="6a725-180">Use only keyboard shortcuts with the following pattern in your add-in: \**Ctrl+Shift+Alt+* x\*\*\*, where *x* is some other key.</span></span>
- <span data-ttu-id="6a725-181">Si vous avez besoin de raccourcis clavier, consultez la liste des [raccourcis](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)clavier Excel et évitez d’en utiliser un dans votre module.</span><span class="sxs-lookup"><span data-stu-id="6a725-181">If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.</span></span>

## <a name="browser-shortcuts-that-cannot-be-overridden"></a><span data-ttu-id="6a725-182">Raccourcis du navigateur qui ne peuvent pas être préférés</span><span class="sxs-lookup"><span data-stu-id="6a725-182">Browser shortcuts that cannot be overridden</span></span>

<span data-ttu-id="6a725-183">Vous ne pouvez utiliser aucune des combinaisons de clavier suivantes.</span><span class="sxs-lookup"><span data-stu-id="6a725-183">You cannot use any of the following keyboard combinations.</span></span> <span data-ttu-id="6a725-184">Ils sont utilisés par les navigateurs et ne peuvent pas être utilisés.</span><span class="sxs-lookup"><span data-stu-id="6a725-184">They are used by browsers and cannot be overridden.</span></span> <span data-ttu-id="6a725-185">Cette liste est un travail en cours.</span><span class="sxs-lookup"><span data-stu-id="6a725-185">This list is a work in progress.</span></span> <span data-ttu-id="6a725-186">Si vous découvrez d’autres combinaisons qui ne peuvent pas être overridées, faites-le nous savoir à l’aide de l’outil de commentaires en bas de cette page.</span><span class="sxs-lookup"><span data-stu-id="6a725-186">If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.</span></span>

- <span data-ttu-id="6a725-187">Ctrl+N</span><span class="sxs-lookup"><span data-stu-id="6a725-187">Ctrl+N</span></span>
- <span data-ttu-id="6a725-188">Ctrl+Shift+N</span><span class="sxs-lookup"><span data-stu-id="6a725-188">Ctrl+Shift+N</span></span>
- <span data-ttu-id="6a725-189">Ctrl+T</span><span class="sxs-lookup"><span data-stu-id="6a725-189">Ctrl+T</span></span>
- <span data-ttu-id="6a725-190">Ctrl+Shift+T</span><span class="sxs-lookup"><span data-stu-id="6a725-190">Ctrl+Shift+T</span></span>
- <span data-ttu-id="6a725-191">Ctrl+W</span><span class="sxs-lookup"><span data-stu-id="6a725-191">Ctrl+W</span></span>
- <span data-ttu-id="6a725-192">Ctrl+PgUp/PgDn</span><span class="sxs-lookup"><span data-stu-id="6a725-192">Ctrl+PgUp/PgDn</span></span>

## <a name="localize-the-keyboard-shortcuts-json"></a><span data-ttu-id="6a725-193">Localisez les raccourcis clavier JSON</span><span class="sxs-lookup"><span data-stu-id="6a725-193">Localize the keyboard shortcuts JSON</span></span>

<span data-ttu-id="6a725-194">Si votre add-in prend en charge plusieurs paramètres régionaux, vous devez trouver la propriété des `name` objets d’action.</span><span class="sxs-lookup"><span data-stu-id="6a725-194">If your add-in supports multiple locales, you'll need to localize the `name` property of the action objects.</span></span> <span data-ttu-id="6a725-195">En outre, si l’un des paramètres régionaux que le add-in prend en charge a des alphabets ou des systèmes d’écriture différents, et par conséquent différents claviers, vous devrez peut-être également trouver les raccourcis.</span><span class="sxs-lookup"><span data-stu-id="6a725-195">Also, if any of the locales that the add-in supports have alphabets or different writing systems, and hence different keyboards, you may need to localize the shortcuts also.</span></span> <span data-ttu-id="6a725-196">Pour plus d’informations sur la façon de trouver les raccourcis clavier JSON, voir [Localize extended overrides](../develop/localization.md#localize-extended-overrides).</span><span class="sxs-lookup"><span data-stu-id="6a725-196">For information about how to localize the keyboard shortcuts JSON, see [Localize extended overrides](../develop/localization.md#localize-extended-overrides).</span></span>

## <a name="next-steps"></a><span data-ttu-id="6a725-197">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="6a725-197">Next Steps</span></span>

- <span data-ttu-id="6a725-198">Consultez l’exemple de raccourcis [clavier-excel pour le add-in.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)</span><span class="sxs-lookup"><span data-stu-id="6a725-198">See the sample add-in [excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>
- <span data-ttu-id="6a725-199">Obtenez une vue d’ensemble de l’utilisation des substitutions étendues dans [Work avec des substitutions étendues du manifeste.](../develop/extended-overrides.md)</span><span class="sxs-lookup"><span data-stu-id="6a725-199">Get an overview of working with extended overrides in [Work with extended overrides of the manifest](../develop/extended-overrides.md).</span></span>
