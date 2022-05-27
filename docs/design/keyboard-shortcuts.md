---
title: Raccourcis clavier personnalisés dans les compléments Office
description: Découvrez comment ajouter des raccourcis clavier personnalisés, également appelés combinaisons de touches, à votre complément Office.
ms.date: 11/22/2021
localization_priority: Normal
ms.openlocfilehash: bd3131ea8e5f0c2f1caadca58ab2e47f588fbfc6
ms.sourcegitcommit: 690c1cc5f9027fd9859e650f3330801fe45e6e67
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/27/2022
ms.locfileid: "65752868"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins"></a>Ajouter des raccourcis clavier personnalisés à vos compléments Office

Les raccourcis clavier, également appelés combinaisons de touches, permettent aux utilisateurs de votre complément de travailler plus efficacement. Les raccourcis clavier améliorent également l’accessibilité du complément pour les utilisateurs handicapés en fournissant une alternative à la souris.

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> Pour commencer avec une version opérationnelle d’un complément avec des raccourcis clavier déjà activés, clonez et exécutez l’exemple [Excel Raccourcis clavier](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts). Lorsque vous êtes prêt à ajouter des raccourcis clavier à votre propre complément, passez à cet article.

Il existe trois étapes pour ajouter des raccourcis clavier à un complément.

1. [Configurez le manifeste du complément](#configure-the-manifest).
1. [Créez ou modifiez le fichier JSON de raccourcis](#create-or-edit-the-shortcuts-json-file) pour définir les actions et leurs raccourcis clavier.
1. [Ajoutez un ou plusieurs appels d’exécution](#create-a-mapping-of-actions-to-their-functions) de l’API [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member) pour mapper une fonction à chaque action.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Il y a deux petites modifications à apporter au manifeste. L’un consiste à permettre au complément d’utiliser un runtime partagé et l’autre à pointer vers un fichier au format JSON où vous avez défini les raccourcis clavier.

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurer le complément pour utiliser un runtime partagé

L’ajout de raccourcis clavier personnalisés nécessite que votre complément utilise le runtime partagé. Pour plus d’informations, [configurez un complément pour utiliser un runtime partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

### <a name="link-the-mapping-file-to-the-manifest"></a>Lier le fichier de mappage au manifeste

Immédiatement *en dessous* (pas à l’intérieur) de l’élément `<VersionOverrides>` dans le manifeste, ajoutez un élément [ExtendedOverrides](/javascript/api/manifest/extendedoverrides) . Définissez l’attribut `Url` sur l’URL complète d’un fichier JSON dans votre projet que vous allez créer à une étape ultérieure.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>Créer ou modifier le fichier JSON de raccourcis

Créez un fichier JSON dans votre projet. Assurez-vous que le chemin d’accès du fichier correspond à l’emplacement que vous avez spécifié pour l’attribut `Url` de l’élément [ExtendedOverrides](/javascript/api/manifest/extendedoverrides) . Ce fichier décrit vos raccourcis clavier et les actions qu’ils appelleront.

1. Dans le fichier JSON, il existe deux tableaux. Le tableau d’actions contient des objets qui définissent les actions à appeler et le tableau de raccourcis contient des objets qui mappent des combinaisons de touches sur des actions. Voici un exemple.
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

    Pour plus d’informations sur les objets JSON, consultez [Construire les objets d’action](#construct-the-action-objects) et [Construire les objets de raccourci](#construct-the-shortcut-objects). Le schéma complet pour les raccourcis JSON se trouve à [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

    > [!NOTE]
    > Vous pouvez utiliser « CONTROL » à la place de « Ctrl » dans cet article.

    Dans une étape ultérieure, les actions seront elles-mêmes mappées aux fonctions que vous écrivez. Dans cet exemple, vous allez ensuite mapper SHOWTASKPANE à une fonction qui appelle la `Office.addin.showAsTaskpane` méthode et HIDETASKPANE à une fonction qui appelle la `Office.addin.hide` méthode.

## <a name="create-a-mapping-of-actions-to-their-functions"></a>Créer un mappage d’actions à leurs fonctions

1. Dans votre projet, ouvrez le fichier JavaScript chargé par votre page HTML dans l’élément `<FunctionFile>` .
1. Dans le fichier JavaScript, utilisez l’API [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member) pour mapper chaque action que vous avez spécifiée dans le fichier JSON à une fonction JavaScript. Ajoutez le Code JavaScript suivant au fichier. Notez ce qui suit concernant le code.

    - Le premier paramètre est l’une des actions du fichier JSON.
    - Le deuxième paramètre est la fonction qui s’exécute lorsqu’un utilisateur appuie sur la combinaison de touches mappée à l’action dans le fichier JSON.

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. Pour continuer l’exemple, utilisez-le `'SHOWTASKPANE'` comme premier paramètre.
1. Pour le corps de la fonction, utilisez la méthode [Office.addin.showAsTaskpane](/javascript/api/office/office.addin#office-office-addin-showastaskpane-member(1)) pour ouvrir le volet Office du complément. Lorsque vous avez terminé, le code doit ressembler à ce qui suit :

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

1. Ajoutez un deuxième appel de `Office.actions.associate` fonction pour mapper l’action `HIDETASKPANE` à une fonction qui appelle [Office.addin.hide](/javascript/api/office/office.addin#office-office-addin-hide-member(1)). Voici un exemple.

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

En suivant les étapes précédentes, votre complément peut activer la visibilité du volet Office en appuyant sur **Ctrl+Alt+Haut** et **Ctrl+Alt+Bas**. Le même comportement s’affiche dans l’exemple [de raccourcis clavier Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts) dans le dépôt PnP des compléments Office dans GitHub.

## <a name="details-and-restrictions"></a>Détails et restrictions

### <a name="construct-the-action-objects"></a>Construire les objets d’action

Utilisez les instructions suivantes lors de la spécification des objets dans le `actions` tableau du fichier shortcuts.json.

- Les noms `id` de propriétés sont `name` obligatoires.
- La `id` propriété est utilisée pour identifier de façon unique l’action à appeler à l’aide d’un raccourci clavier.
- La `name` propriété doit être une chaîne conviviale décrivant l’action. Il doit s’agir d’une combinaison des caractères A - Z, a - z, 0 - 9, et les signes de ponctuation « - », « _ » et « + ».
- La propriété `type` est facultative. Actuellement, seul `ExecuteFunction` le type est pris en charge.

Voici un exemple.

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

Le schéma complet pour les raccourcis JSON se trouve à [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

### <a name="construct-the-shortcut-objects"></a>Construire les objets de raccourci

Utilisez les instructions suivantes lors de la spécification des objets dans le `shortcuts` tableau du fichier shortcuts.json.

- Les noms de `action`propriétés, `key`et `default` sont obligatoires.
- La valeur de la `action` propriété est une chaîne et doit correspondre à l’une `id` des propriétés de l’objet d’action.
- La `default` propriété peut être n’importe quelle combinaison des caractères A - Z, a -z, 0 - 9, et les signes de ponctuation « - », « _ » et « + ». (Par convention, les lettres minuscules ne sont pas utilisées dans ces propriétés.)
- La `default` propriété doit contenir le nom d’au moins une clé de modificateur (Alt, Ctrl, Maj) et une seule autre clé.
- La touche Maj ne peut pas être utilisée comme seule touche modificative. Combinez Maj avec Alt ou Ctrl.
- Pour les Mac, nous prenons également en charge la touche modificateur de commande.
- Pour les Mac, Alt est mappé à la clé Option. Pour Windows, la commande est mappée à la touche Ctrl.
- Lorsque deux caractères sont liés à la même clé physique dans un clavier standard, ils sont synonymes dans la `default` propriété ; par exemple, Alt+a et Alt+A sont le même raccourci, de même que Ctrl+- et Ctrl+\_ , car « - » et « _ » sont la même clé physique.
- Le caractère « + » indique que les touches de chaque côté de celui-ci sont enfoncées simultanément.

Voici un exemple.

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

Le schéma complet pour les raccourcis JSON se trouve à [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

> [!NOTE]
> Les touches d’accès, également appelées raccourcis clés séquentiels, tels que le raccourci Excel pour choisir une couleur de remplissage **Alt+H, H**, ne sont pas prises en charge dans Office compléments.

## <a name="avoid-key-combinations-in-use-by-other-add-ins"></a>Éviter les combinaisons de touches utilisées par d’autres compléments

De nombreux raccourcis clavier sont déjà utilisés par Office. Évitez d’inscrire des raccourcis clavier pour votre complément qui sont déjà utilisés. Toutefois, il peut y avoir certaines instances où il est nécessaire de remplacer les raccourcis clavier existants ou de gérer les conflits entre plusieurs compléments qui ont inscrit le même raccourci clavier.

En cas de conflit, l’utilisateur voit une boîte de dialogue la première fois qu’il tente d’utiliser un raccourci clavier en conflit. Notez que le texte de l’option de complément qui s’affiche dans cette boîte de dialogue provient de la `name` propriété de l’objet d’action dans le `shortcuts.json` fichier.

![Illustration montrant un conflit modal avec deux actions différentes pour un raccourci unique.](../images/add-in-shortcut-conflict-modal.png)

L’utilisateur peut sélectionner l’action que le raccourci clavier effectuera. Une fois la sélection effectuée, la préférence est enregistrée pour les utilisations futures du même raccourci. Les préférences de raccourci sont enregistrées par utilisateur, par plateforme. Si l’utilisateur souhaite modifier ses préférences, il peut appeler la commande **Réinitialiser les préférences de raccourci des compléments Office** à partir de la zone de recherche **Rechercher**. L’appel de la commande efface toutes les préférences de raccourci de complément de l’utilisateur et l’utilisateur est à nouveau invité à entrer la boîte de dialogue de conflit la prochaine fois qu’il tente d’utiliser un raccourci en conflit.

![La zone de recherche Rechercher dans Excel montrant l’action de réinitialisation Office les préférences de raccourci de complément.](../images/add-in-reset-shortcuts-action.png)

Pour une expérience utilisateur optimale, nous vous recommandons de réduire les conflits avec Excel avec ces bonnes pratiques.

- Utilisez uniquement les raccourcis clavier avec le modèle suivant : **Ctrl+Maj+Alt+* x***, où *x* est une autre clé.
- Si vous avez besoin de raccourcis clavier supplémentaires, consultez la [liste des raccourcis clavier Excel](https://support.microsoft.com/office/1798d9d5-842a-42b8-9c99-9b7213f0040f) et évitez d’en utiliser un dans votre complément.
- Lorsque le focus clavier se trouve à l’intérieur de l’interface utilisateur du complément, **Ctrl+Espace** et **Ctrl+Maj+F10** ne fonctionnent pas, car il s’agit de raccourcis d’accessibilité essentiels.
- Sur un ordinateur Windows ou Mac, si la commande « Réinitialiser les préférences de raccourci des compléments Office » n’est pas disponible dans le menu de recherche, l’utilisateur peut ajouter manuellement la commande au ruban en personnalisant le ruban via le menu contextuel.

## <a name="customize-the-keyboard-shortcuts-per-platform"></a>Personnaliser les raccourcis clavier par plateforme

Il est possible de personnaliser les raccourcis pour qu’ils soient spécifiques à la plateforme. Voici un exemple de l’objet `shortcuts` qui personnalise les raccourcis pour chacune des plateformes suivantes : `windows`, `web``mac`. Notez que vous devez toujours avoir une `default` touche de raccourci pour chaque raccourci.

Dans l’exemple suivant, la `default` clé est la clé de secours pour toute plateforme qui n’est pas spécifiée. La seule plateforme non spécifiée étant Windows, la `default` clé s’applique uniquement à Windows.

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

## <a name="localize-the-keyboard-shortcuts-json"></a>Localiser les raccourcis clavier JSON

Si votre complément prend en charge plusieurs paramètres régionaux, vous devez localiser la `name` propriété des objets d’action. En outre, si l’un des paramètres régionaux pris en charge par le complément a des alphabets ou des systèmes d’écriture différents, et donc des claviers différents, vous devrez peut-être également localiser les raccourcis. Pour plus d’informations sur la localisation des raccourcis clavier JSON, consultez [Localiser les remplacements étendus](../develop/localization.md#localize-extended-overrides).

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>Raccourcis du navigateur qui ne peuvent pas être remplacés

Lorsque vous utilisez des raccourcis clavier personnalisés sur le web, certains raccourcis clavier utilisés par le navigateur ne peuvent pas être remplacés par des compléments. Cette liste est un travail en cours. Si vous découvrez d’autres combinaisons qui ne peuvent pas être remplacées, faites-le nous savoir à l’aide de l’outil de commentaires en bas de cette page.

- Ctrl+N
- Ctrl+Maj+N
- Ctrl+T
- Ctrl+Maj+T
- Ctrl+W
- Ctrl+PgUp/PgDn

## <a name="enable-custom-keyboard-shortcuts-for-specific-users"></a>Activer des raccourcis clavier personnalisés pour des utilisateurs spécifiques

Votre complément peut permettre aux utilisateurs de réaffecter les actions du complément à d’autres combinaisons de clavier.

> [!NOTE]
> Les API décrites dans cette section nécessitent l’ensemble de conditions requises [KeyboardShortcuts 1.1](/javascript/api/requirement-sets/common/keyboard-shortcuts-requirement-sets) .

Utilisez la méthode [Office.actions.replaceShortcuts](/javascript/api/office/office.actions#office-office-actions-replaceshortcuts-member) pour affecter les combinaisons de clavier personnalisées d’un utilisateur à vos actions de compléments. La méthode prend un paramètre de type `{[actionId:string]: string|null}`, où les `actionId`S sont un sous-ensemble des ID d’action qui doivent être définis dans le code JSON du manifeste étendu du complément. Les valeurs sont les combinaisons de touches préférées de l’utilisateur. La valeur peut également être `null`, ce qui supprimera toute personnalisation pour cela `actionId` et rétablira la combinaison de clavier par défaut définie dans le code JSON du manifeste étendu du complément.

Si l’utilisateur est connecté à Office, les combinaisons personnalisées sont enregistrées dans les paramètres d’itinérance de l’utilisateur par plateforme. La personnalisation des raccourcis n’est actuellement pas prise en charge pour les utilisateurs anonymes.

```javascript
const userCustomShortcuts = {
    SHOWTASKPANE:"CTRL+SHIFT+1", 
    HIDETASKPANE:"CTRL+SHIFT+2"
};
Office.actions.replaceShortcuts(userCustomShortcuts)
    .then(function () {
        console.log("Successfully registered.");
    })
    .catch(function (ex) {
        if (ex.code == "InvalidOperation") {
            console.log("ActionId does not exist or shortcut combination is invalid.");
        }
    });
```

Pour savoir quels raccourcis sont déjà utilisés pour l’utilisateur, appelez la méthode [Office.actions.getShortcuts](/javascript/api/office/office.actions#office-office-actions-getshortcuts-member). Cette méthode retourne un objet de type `[actionId:string]:string|null}`, où les valeurs représentent la combinaison de clavier actuelle que l’utilisateur doit utiliser pour appeler l’action spécifiée. Les valeurs peuvent provenir de trois sources différentes :

- En cas de conflit avec le raccourci et que l’utilisateur a choisi d’utiliser une autre action (native ou un autre complément) pour cette combinaison de clavier, la valeur retournée est `null` puisque le raccourci a été remplacé et qu’il n’existe aucune combinaison de clavier que l’utilisateur peut actuellement utiliser pour appeler cette action de complément.
- Si le raccourci a été personnalisé à l’aide de la méthode [Office.actions.replaceShortcuts](/javascript/api/office/office.actions#office-office-actions-replaceshortcuts-member), la valeur retournée est la combinaison de clavier personnalisée.
- Si le raccourci n’a pas été remplacé ou personnalisé, il retourne la valeur du code JSON du manifeste étendu du complément.

Voici un exemple.

```javascript
Office.actions.getShortcuts()
    .then(function (userShortcuts) {
       for (const action in userShortcuts) {
           let shortcut = userShortcuts[action];
           console.log(action + ": " + shortcut);
       }
    });

```

Comme décrit dans [Éviter les combinaisons de touches utilisées par d’autres compléments](#avoid-key-combinations-in-use-by-other-add-ins), il est recommandé d’éviter les conflits dans les raccourcis. Pour déterminer si une ou plusieurs combinaisons de clés sont déjà utilisées, passez-les en tant que tableau de chaînes à la méthode [Office.actions.areShortcutsInUse](/javascript/api/office/office.actions#office-office-actions-areshortcutsinuse-member). La méthode retourne un rapport contenant des combinaisons de clés déjà utilisées sous la forme d’un tableau d’objets de type `{shortcut: string, inUse: boolean}`. La `shortcut` propriété est une combinaison de touches, telle que « Ctrl+Maj+1 ». Si la combinaison est déjà inscrite dans une autre action, la `inUse` propriété est définie sur `true`. Par exemple : `[{shortcut: "CTRL+SHIFT+1", inUse: true}, {shortcut: "CTRL+SHIFT+2", inUse: false}]`. L’extrait de code suivant est un exemple :

```javascript
const shortcuts = ["CTRL+SHIFT+1", "CTRL+SHIFT+2"];
Office.actions.areShortcutsInUse(shortcuts)
    .then(function (inUseArray) {
        const availableShortcuts = inUseArray.filter(function (shortcut) { return !shortcut.inUse; });
        console.log(availableShortcuts);
        const usedShortcuts = inUseArray.filter(function (shortcut) { return shortcut.inUse; });
        console.log(usedShortcuts);
    });

```

## <a name="next-steps"></a>Étapes suivantes

- Consultez [l’exemple de complément Excel raccourcis clavier](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts).
- Obtenez une vue d’ensemble de l’utilisation des remplacements étendus dans [Work avec des remplacements étendus du manifeste](../develop/extended-overrides.md).
