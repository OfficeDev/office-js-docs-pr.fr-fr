---
title: Raccourcis clavier personnalisés dans Office des modules
description: 'Découvrez comment ajouter des raccourcis clavier personnalisés, également appelés combinaisons de touches, à votre Office de clavier.'
ms.date: 11/22/2021
localization_priority: Normal
---

# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins"></a>Ajouter des raccourcis clavier personnalisés à vos Office de travail

Les raccourcis clavier, également appelés combinaisons de touches, permettent aux utilisateurs de votre module de travailler plus efficacement. Les raccourcis clavier améliorent également l’accessibilité du module pour les utilisateurs présentant un handicap en offrant une alternative à la souris.

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> Pour commencer avec une version de travail d’un add-in avec des raccourcis clavier déjà activés, clonez et exécutez l’exemple Excel [raccourcis clavier](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts). Lorsque vous êtes prêt à ajouter des raccourcis clavier à votre propre add-in, poursuivez avec cet article.

Il existe trois étapes pour ajouter des raccourcis clavier à un module.

1. [Configurez le manifeste du module.](#configure-the-manifest)
1. [Créez ou modifiez le fichier JSON de raccourcis](#create-or-edit-the-shortcuts-json-file) pour définir des actions et leurs raccourcis clavier.
1. [Ajoutez un ou plusieurs appels runtime](#create-a-mapping-of-actions-to-their-functions) de [l’API Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member) pour ma cartographier une fonction à chaque action.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Deux petites modifications sont à apporter au manifeste. L’une consiste à permettre au add-in d’utiliser un runtime partagé et l’autre à pointer vers un fichier au format JSON où vous avez défini les raccourcis clavier.

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurer le add-in pour utiliser un runtime partagé

L’ajout de raccourcis clavier personnalisés nécessite que votre add-in utilise le runtime partagé. Pour plus d’informations, [configurez un module complémentaire pour utiliser un runtime partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

### <a name="link-the-mapping-file-to-the-manifest"></a>Lier le fichier de mappage au manifeste

Juste *en dessous* (pas à l’intérieur) `<VersionOverrides>` de l’élément dans le manifeste, ajoutez [un élément ExtendedOverrides](../reference/manifest/extendedoverrides.md) . Définissez l’attribut `Url` sur l’URL complète d’un fichier JSON dans votre projet que vous créerez à une étape ultérieure.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>Créer ou modifier le fichier JSON de raccourcis

Créez un fichier JSON dans votre projet. Assurez-vous que le chemin d’accès au fichier correspond `Url` à l’emplacement que vous avez spécifié pour l’attribut de l’élément [ExtendedOverrides](../reference/manifest/extendedoverrides.md) . Ce fichier décrit vos raccourcis clavier et les actions qu’ils appelleront.

1. Le fichier JSON se trouve à l’intérieur de deux tableaux. Le tableau d’actions contient des objets qui définissent les actions à appeler et le tableau de raccourcis contient des objets qui maient des combinaisons de touches sur des actions. Voici un exemple.
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

    Pour plus d’informations sur les objets JSON, voir [Construct the action objects](#construct-the-action-objects) and [Construct the shortcut objects](#construct-the-shortcut-objects). Le schéma complet des raccourcis JSON se trouve à [l’extension-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

    > [!NOTE]
    > Vous pouvez utiliser « CONTROL » à la place de « Ctrl » tout au long de cet article.

    Dans une étape ultérieure, les actions seront elles-mêmes mappées aux fonctions que vous écrivez. Dans cet exemple, vous masquez ultérieurement SHOWTASKPANE `Office.addin.showAsTaskpane` à une fonction qui appelle la méthode et HIDETASKPANE à une fonction qui appelle la `Office.addin.hide` méthode.

## <a name="create-a-mapping-of-actions-to-their-functions"></a>Créer un mappage des actions à leurs fonctions

1. Dans votre projet, ouvrez le fichier JavaScript chargé par votre page HTML dans l’élément `<FunctionFile>` .
1. Dans le fichier JavaScript, utilisez l’API [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member) pour ma cartographier chaque action que vous avez spécifiée dans le fichier JSON sur une fonction JavaScript. Ajoutez le javaScript suivant au fichier. Notez ce qui suit à propos du code.

    - Le premier paramètre est l’une des actions du fichier JSON.
    - Le deuxième paramètre est la fonction qui s’exécute lorsqu’un utilisateur appuie sur la combinaison de touches mappée à l’action dans le fichier JSON.

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. Pour continuer l’exemple, utilisez-le `'SHOWTASKPANE'` comme premier paramètre.
1. Pour le corps de la fonction, utilisez la [méthode Office.addin.showAsTaskpane](/javascript/api/office/office.addin#office-office-addin-showastaskpane-member(1)) pour ouvrir le volet Des tâches du module. Lorsque vous avez terminé, le code doit ressembler à ce qui suit :

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

1. Ajoutez un deuxième appel de `Office.actions.associate` fonction pour maque l’action `HIDETASKPANE` à une fonction qui appelle [Office.addin.hide](/javascript/api/office/office.addin#office-office-addin-hide-member(1)). Voici un exemple.

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

La suite des étapes précédentes permet à votre add-in de faire tourner la visibilité du volet Des tâches en appuyant sur **Ctrl+Alt+Haut** et **Ctrl+Alt+Bas**. Le même comportement est illustré dans [l’exemple de raccourcis clavier Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts) dans le Office PnP des GitHub.

## <a name="details-and-restrictions"></a>Détails et restrictions

### <a name="construct-the-action-objects"></a>Construire les objets d’action

Utilisez les instructions suivantes lors de la spécification des objets dans le `actions` tableau du shortcuts.json.

- Les noms des propriétés `id` `name` sont obligatoires.
- La `id` propriété est utilisée pour identifier de manière unique l’action à appeler à l’aide d’un raccourci clavier.
- La `name` propriété doit être une chaîne conviviale décrivant l’action. Il doit s’agit d’une combinaison des caractères A - Z, a - z, 0 - 9, et des signes de ponctuation « - », « _ » et « + ».
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

Le schéma complet des raccourcis JSON se trouve à [l’extension-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

### <a name="construct-the-shortcut-objects"></a>Construire les objets de raccourci

Utilisez les instructions suivantes lors de la spécification des objets dans le `shortcuts` tableau du shortcuts.json.

- Les noms des propriétés `action`et `key`sont `default` obligatoires.
- La valeur de la propriété `action` est une chaîne et doit correspondre à l’une des propriétés `id` de l’objet action.
- La `default` propriété peut être n’importe quelle combinaison des caractères A - Z, -z, 0 - 9 et les signes de ponctuation « - », « _ » et « + ». (Par convention, les lettres majuscules ne sont pas utilisées dans ces propriétés.)
- La `default` propriété doit contenir le nom d’au moins une touche de modification (Alt, Ctrl, Shift) et une seule autre touche.
- Shift ne peut pas être utilisé comme seule touche de modification. Combinez Shift avec Alt ou Ctrl.
- Pour les Mac, nous  pris en charge également la touche Modificateur de commande.
- Pour les Mac, Alt est mappé à la touche Option. Pour Windows, La commande est mappée sur la touche Ctrl.
- Lorsque deux caractères sont liés à la même touche physique dans un clavier standard, ils sont synonymes `default` dans la propriété ; par exemple, Alt+a et Alt+A sont le même raccourci, tout comme Ctrl+- et Ctrl+\_ car « - » et « _ » sont la même touche physique.
- Le caractère « + » indique que les touches de chaque côté de celui-ci sont entrées simultanément.

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

Le schéma complet des raccourcis JSON se trouve à [l’extension-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

> [!NOTE]
> Les touches d’accès, également appelées raccourcis de touches séquentiels, tels que le raccourci Excel pour choisir une couleur de remplissage **Alt+H, H**, ne sont pas pris en charge dans les Office.

## <a name="avoid-key-combinations-in-use-by-other-add-ins"></a>Éviter les combinaisons de touches en cours d’utilisation par d’autres modules

De nombreux raccourcis clavier sont déjà utilisés par les Office. Évitez d’inscrire des raccourcis clavier pour votre module qui sont déjà utilisés. Cependant, dans certains cas, il peut être nécessaire de remplacer les raccourcis clavier existants ou de gérer les conflits entre plusieurs modules qui ont inscrit le même raccourci clavier.

En cas de conflit, l’utilisateur voit une boîte de dialogue la première fois qu’il tente d’utiliser un raccourci clavier en conflit. Notez que le texte de l’option `name` de module qui s’affiche dans cette boîte de dialogue provient de la propriété de l’objet action dans le `shortcuts.json` fichier.

![Illustration montrant un conflit modal avec deux actions différentes pour un seul raccourci.](../images/add-in-shortcut-conflict-modal.png)

L’utilisateur peut sélectionner l’action que le raccourci clavier va prendre. Après avoir fait la sélection, la préférence est enregistrée pour les futures utilisations du même raccourci. Les préférences de raccourci sont enregistrées par utilisateur, par plateforme. Si l’utilisateur souhaite modifier ses préférences, il peut appeler la commande  Réinitialiser les préférences de raccourci des Office dans la zone de recherche Rechercher. L’appel de la commande permet d’effacer toutes les préférences de raccourci de l’utilisateur et l’utilisateur sera de nouveau invité à utiliser la boîte de dialogue de conflit la prochaine fois qu’il tentera d’utiliser un raccourci conflictuelle.

![La zone de recherche Rechercher dans Excel affiche la réinitialisation Office’action de préférence de raccourci de l’ajout.](../images/add-in-reset-shortcuts-action.png)

Pour une expérience utilisateur de qualité, nous vous recommandons de minimiser les conflits Excel avec ces bonnes pratiques.

- Utilisez uniquement les raccourcis clavier avec le modèle suivant : **Ctrl+Shift+Alt+* x***, où *x* est une autre touche.
- Si vous avez besoin de raccourcis clavier, consultez la liste des [raccourcis](https://support.microsoft.com/office/1798d9d5-842a-42b8-9c99-9b7213f0040f) clavier Excel et évitez d’en utiliser dans votre module.
- Lorsque le focus du clavier se trouve à l’intérieur de l’interface utilisateur du module, **Ctrl+Espace et** **Ctrl+Shift+F10** ne fonctionnent pas, car il s’agit de raccourcis d’accessibilité essentiels.
- Sur un ordinateur Windows ou Mac, si la commande « Réinitialiser les préférences de raccourci des macros de Office » n’est pas disponible dans le menu de recherche, l’utilisateur peut ajouter manuellement la commande au ruban en personnalisant le ruban via le menu contexté.

## <a name="customize-the-keyboard-shortcuts-per-platform"></a>Personnaliser les raccourcis clavier par plateforme

Il est possible de personnaliser les raccourcis pour qu’ils soient spécifiques à la plateforme. Voici un exemple de l’objet `shortcuts` qui personnalise les raccourcis pour chacune des plateformes suivantes : `windows`, , `mac``web`. Notez que vous devez toujours avoir une touche `default` de raccourci pour chaque raccourci.

Dans l’exemple suivant, la `default` clé est la clé de retour pour toute plateforme qui n’est pas spécifiée. La seule plateforme non spécifiée est Windows, `default` donc la clé s’applique uniquement aux Windows.

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

## <a name="localize-the-keyboard-shortcuts-json"></a>Localisez les raccourcis clavier JSON

Si votre add-in prend en charge plusieurs paramètres régionaux, vous devez trouver `name` la propriété des objets d’action. En outre, si l’un des paramètres régionaux que le add-in prend en charge a des alphabets ou des systèmes d’écriture différents, et par conséquent différents claviers, vous devrez peut-être également trouver les raccourcis. Pour plus d’informations sur la façon de trouver les raccourcis clavier JSON, voir [Localize extended overrides](../develop/localization.md#localize-extended-overrides).

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>Raccourcis du navigateur qui ne peuvent pas être préférés

Lorsque vous utilisez des raccourcis clavier personnalisés sur le web, certains raccourcis clavier utilisés par le navigateur ne peuvent pas être préférés par les modules. Cette liste est un travail en cours. Si vous découvrez d’autres combinaisons qui ne peuvent pas être overridées, faites-le nous savoir à l’aide de l’outil de commentaires en bas de cette page.

- Ctrl+N
- Ctrl+Shift+N
- Ctrl+T
- Ctrl+Shift+T
- Ctrl+W
- Ctrl+PgUp/PgDn

## <a name="enable-custom-keyboard-shortcuts-for-specific-users-preview"></a>Activer les raccourcis clavier personnalisés pour des utilisateurs spécifiques (aperçu)

Votre add-in peut permettre aux utilisateurs de réaffecter les actions du module à d’autres combinaisons de clavier.

> [!IMPORTANT]
> Les fonctionnalités décrites dans cette section sont actuellement en prévisualisation et peuvent faire l’objet de changements. Elles ne sont pas prises en charge dans les environnements de production pour l’instant. Pour essayer les fonctionnalités d’aperçu, vous devez rejoindre le [programme Office Insider](https://insider.office.com/join).
> Un bon moyen de tester les fonctionnalités en préversion consiste à utiliser un abonnement Microsoft 365. Si vous n’avez pas déjà d’abonnement Microsoft 365, vous pouvez en obtenir un gratuitement en rejoignant le [Programme pour les développeurs Microsoft 365](https://developer.microsoft.com/office/dev-program).

> [!NOTE]
> Les API décrites dans cette section nécessitent [l’ensemble de conditions requises KeyboardShortcuts 1.1](../reference/requirement-sets/keyboard-shortcuts-requirement-sets.md) .

Utilisez la [Office.actions.replaceShortcuts](/javascript/api/office/office.actions#office-office-actions-replaceshortcuts-member) pour affecter les combinaisons de clavier personnalisées d’un utilisateur à vos actions de modules. La méthode prend un paramètre de type `{[actionId:string]: string|null}`, `actionId`où les s sont un sous-ensemble des ID d’action qui doivent être définis dans le manifeste JSON étendu du module. Les valeurs sont les combinaisons de touches préférées de l’utilisateur. La valeur peut également `null`être , `actionId` ce qui permet de supprimer toute personnalisation pour cela et de revenir à la combinaison de clavier par défaut définie dans le manifeste JSON étendu du module.

Si l’utilisateur est connecté Office, les combinaisons personnalisées sont enregistrées dans les paramètres d’itinérance de l’utilisateur par plateforme. La personnalisation des raccourcis n’est actuellement pas prise en charge pour les utilisateurs anonymes.

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

Pour savoir quels raccourcis sont déjà utilisés pour l’utilisateur, appelez la [méthode Office.actions.getShortcuts](/javascript/api/office/office.actions#office-office-actions-getshortcuts-member). Cette méthode renvoie un objet de type `[actionId:string]:string|null}`, où les valeurs représentent la combinaison de clavier actuelle que l’utilisateur doit utiliser pour appeler l’action spécifiée. Les valeurs peuvent être provenant de trois sources différentes :

- S’il y a eu un conflit avec le raccourci et que l’utilisateur a choisi d’utiliser une autre action (native ou autre) pour cette combinaison de clavier, `null` la valeur renvoyée sera puisque le raccourci a été changé et qu’il n’existe aucune combinaison de clavier que l’utilisateur peut utiliser actuellement pour appeler cette action de module.
- Si le raccourci a été personnalisé à l’aide de [la méthode Office.actions.replaceShortcuts](/javascript/api/office/office.actions#office-office-actions-replaceshortcuts-member), la valeur renvoyée sera la combinaison de clavier personnalisée.
- Si le raccourci n’a pas été overrided ou personnalisé, il retourne la valeur à partir du manifeste JSON étendu du module.

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

Comme décrit dans [Éviter les combinaisons de touches](#avoid-key-combinations-in-use-by-other-add-ins) en cours d’utilisation par d’autres modules, il est bon d’éviter les conflits dans les raccourcis. Pour découvrir si une ou plusieurs combinaisons de touches sont déjà utilisées, passez-les en tant que tableau de chaînes à la méthode [Office.actions.areShortcutsInUse](/javascript/api/office/office.actions#office-office-actions-areshortcutsinuse-member). La méthode renvoie un rapport contenant des combinaisons de touches qui sont déjà utilisées sous la forme d’un tableau d’objets de type `{shortcut: string, inUse: boolean}`. La `shortcut` propriété est une combinaison de touches, telle que « Ctrl+Shift+1 ». Si la combinaison est déjà inscrite dans une autre action, la `inUse` propriété est définie `true`sur . Par exemple, `[{shortcut: "CTRL+SHIFT+1", inUse: true}, {shortcut: "CTRL+SHIFT+2", inUse: false}]`. L’extrait de code suivant est un exemple :

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

- Consultez l [Excel exemple de raccourcis](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts) clavier.
- Obtenez une vue d’ensemble de l’utilisation des substitutions étendues dans [Work avec les substitutions étendues du manifeste](../develop/extended-overrides.md).
