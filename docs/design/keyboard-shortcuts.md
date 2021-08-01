---
title: Raccourcis clavier personnalisés dans les Office des modules
description: Découvrez comment ajouter des raccourcis clavier personnalisés, également appelés combinaisons de touches, à votre Office de clavier.
ms.date: 06/02/2021
localization_priority: Normal
ms.openlocfilehash: 75a7de576368e85436b4d97a4561d609b654642e
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671400"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins"></a>Ajouter des raccourcis clavier personnalisés à vos Office de travail

Les raccourcis clavier, également appelés combinaisons de touches, permettent aux utilisateurs de votre module de travailler plus efficacement. Les raccourcis clavier améliorent également l’accessibilité du module pour les utilisateurs présentant un handicap en offrant une alternative à la souris.

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> Pour commencer avec une version de travail d’un add-in avec des raccourcis clavier déjà activés, clonez et exécutez l’exemple [Excel raccourcis clavier.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) Lorsque vous êtes prêt à ajouter des raccourcis clavier à votre propre add-in, poursuivez avec cet article.

Il existe trois étapes pour ajouter des raccourcis clavier à un add-in :

1. [Configurez le manifeste du add-in.](#configure-the-manifest)
1. [Créez ou modifiez le fichier JSON](#create-or-edit-the-shortcuts-json-file) de raccourcis pour définir des actions et leurs raccourcis clavier.
1. [Ajoutez un ou plusieurs appels runtime](#create-a-mapping-of-actions-to-their-functions) de [l’API Office.actions.associate](/javascript/api/office/office.actions#associate) pour ma cartographier une fonction à chaque action.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Deux petites modifications sont à apporter au manifeste. L’une consiste à permettre au add-in d’utiliser un runtime partagé et l’autre consiste à pointer vers un fichier au format JSON où vous avez défini les raccourcis clavier.

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurer le add-in pour utiliser un runtime partagé

L’ajout de raccourcis clavier personnalisés nécessite que votre add-in utilise le runtime partagé. Pour plus d’informations, [configurez un module complémentaire pour utiliser un runtime partagé.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)

### <a name="link-the-mapping-file-to-the-manifest"></a>Lier le fichier de mappage au manifeste

Juste *en dessous* (pas à l’intérieur) de l’élément dans le manifeste, ajoutez un élément `<VersionOverrides>` [ExtendedOverrides.](../reference/manifest/extendedoverrides.md) Définissez l’attribut sur l’URL complète d’un fichier JSON dans votre projet que `Url` vous créerez à une étape ultérieure.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>Créer ou modifier le fichier JSON de raccourcis

Créez un fichier JSON dans votre projet. Assurez-vous que le chemin d’accès au fichier correspond à l’emplacement que vous avez spécifié pour l’attribut de l’élément `Url` [ExtendedOverrides.](../reference/manifest/extendedoverrides.md) Ce fichier décrit vos raccourcis clavier et les actions qu’ils appelleront.

1. Le fichier JSON se trouve à l’intérieur de deux tableaux. Le tableau d’actions contient des objets qui définissent les actions à appeler et le tableau de raccourcis contient des objets qui maient des combinaisons de touches sur des actions. Voici un exemple :

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

    Pour plus d’informations sur les objets JSON, voir [Construct the action objects](#construct-the-action-objects) and [Construct the shortcut objects](#construct-the-shortcut-objects). The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

    > [!NOTE]
    > Vous pouvez utiliser « CONTROL » à la place de « Ctrl » tout au long de cet article.

    Dans une étape ultérieure, les actions seront elles-mêmes mappées sur les fonctions que vous écrivez. Dans cet exemple, vous masquez ultérieurement SHOWTASKPANE à une fonction qui appelle la méthode et HIDETASKPANE à une fonction qui `Office.addin.showAsTaskpane` appelle la `Office.addin.hide` méthode.

## <a name="create-a-mapping-of-actions-to-their-functions"></a>Créer un mappage des actions à leurs fonctions

1. Dans votre projet, ouvrez le fichier JavaScript chargé par votre page HTML dans `<FunctionFile>` l’élément.
1. Dans le fichier JavaScript, utilisez l’API [Office.actions.associate](/javascript/api/office/office.actions#associate) pour ma cartographier chaque action que vous avez spécifiée dans le fichier JSON sur une fonction JavaScript. Ajoutez le javaScript suivant au fichier. Notez ce qui suit à propos du code.

    - Le premier paramètre est l’une des actions du fichier JSON.
    - Le deuxième paramètre est la fonction qui s’exécute lorsqu’un utilisateur appuie sur la combinaison de touches mappée à l’action dans le fichier JSON.

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. Pour continuer l’exemple, `'SHOWTASKPANE'` utilisez-le comme premier paramètre.
1. Pour le corps de la fonction, utilisez la [méthode Office.addin.showTaskpane](/javascript/api/office/office.addin#showAsTaskpane__) pour ouvrir le volet Des tâches du module. Lorsque vous avez terminé, le code doit ressembler à ce qui suit :

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

1. Ajoutez un deuxième appel de fonction pour maque l’action à une `Office.actions.associate` fonction qui appelle `HIDETASKPANE` [Office.addin.hide](/javascript/api/office/office.addin#hide__). Voici un exemple.

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

La suite des étapes précédentes permet à votre add-in de faire tourner la visibilité du volet Des tâches en appuyant sur **Ctrl+Alt+Haut** et **Ctrl+Alt+Bas.** Le même comportement est illustré dans [l’exemple de raccourcis](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) clavier Excel dans le Office PnP des Office dans GitHub.

## <a name="details-and-restrictions"></a>Détails et restrictions

### <a name="construct-the-action-objects"></a>Construire les objets d’action

Utilisez les instructions suivantes lors de la spécification des objets dans le tableau de la `actions` shortcuts.jssur.

- Les noms des `id` propriétés `name` et sont obligatoires.
- La `id` propriété est utilisée pour identifier de manière unique l’action à appeler à l’aide d’un raccourci clavier.
- La `name` propriété doit être une chaîne conviviale décrivant l’action. Il doit s’agit d’une combinaison des caractères A - Z, a - z, 0 - 9 et des signes de ponctuation « - », « _ » et « + ».
- La propriété `type` est facultative. Actuellement, `ExecuteFunction` seul le type est pris en charge.

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

The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

### <a name="construct-the-shortcut-objects"></a>Construire les objets de raccourci

Utilisez les instructions suivantes lors de la spécification des objets dans le tableau de la `shortcuts` shortcuts.jssur.

- Les noms des `action` propriétés `key` et sont `default` obligatoires.
- La valeur de la propriété est une chaîne et doit correspondre à l’une `action` des `id` propriétés de l’objet action.
- La propriété peut être n’importe quelle combinaison des caractères `default` A - Z, -z, 0 - 9 et les signes de ponctuation « - », « _ » et « + ». (Par convention, les lettres majuscules ne sont pas utilisées dans ces propriétés.)
- La propriété doit contenir le nom d’au moins une touche de `default` modification (Alt, Ctrl, Shift) et une seule autre touche. 
- Shift ne peut pas être utilisé comme seule touche de modification. Combinez Shift avec Alt ou Ctrl.
- Pour les Mac, nous prise en charge également la touche Modificateur de commande.
- Pour les Mac, Alt est mappée sur la touche Option. Pour Windows, Command est mappée sur la touche Ctrl.
- Lorsque deux caractères sont liés à la même touche physique dans un clavier standard, ils sont synonymes dans la propriété ; par exemple, Alt+a et Alt+A sont les mêmes raccourcis, c’est le cas de `default` Ctrl+- et Ctrl+ car « - » et « _ » sont la même touche \_ physique.
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

The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

> [!NOTE]
> Les touches d’accès, également appelées raccourcis de touche séquentiels, tels que le raccourci Excel pour choisir une couleur de remplissage **Alt+H, H,** ne sont pas pris en charge dans les Office.

## <a name="avoid-key-combinations-in-use-by-other-add-ins"></a>Éviter les combinaisons de touches en cours d’utilisation par d’autres modules

De nombreux raccourcis clavier sont déjà utilisés par les Office. Évitez d’inscrire des raccourcis clavier pour votre module qui sont déjà utilisés. Cependant, dans certains cas, il peut être nécessaire de remplacer les raccourcis clavier existants ou de gérer les conflits entre plusieurs modules qui ont inscrit le même raccourci clavier.

En cas de conflit, l’utilisateur voit une boîte de dialogue la première fois qu’il tente d’utiliser un raccourci clavier en conflit, notez que le nom de l’action qui s’affiche dans cette boîte de dialogue est la propriété de l’objet action dans le `name` `shortcuts.json` fichier.

![Illustration montrant un conflit modal avec deux actions différentes pour un seul raccourci.](../images/add-in-shortcut-conflict-modal.png)

L’utilisateur peut sélectionner l’action que le raccourci clavier va prendre. Après avoir fait la sélection, la préférence est enregistrée pour les futures utilisations du même raccourci. Les préférences de raccourci sont enregistrées par utilisateur, par plateforme. Si l’utilisateur souhaite modifier ses préférences,  il peut appeler la commande Réinitialiser  les préférences de raccourci des Office dans la zone de recherche Rechercher. L’appel de la commande permet d’effacer toutes les préférences de raccourci de l’utilisateur et l’utilisateur sera de nouveau invité à utiliser la boîte de dialogue de conflit la prochaine fois qu’il tentera d’utiliser un raccourci conflictuelle :

![La zone de recherche Rechercher dans Excel affiche la réinitialisation Office l’action de préférence de raccourci de l’ajout.](../images/add-in-reset-shortcuts-action.png)

Pour une expérience utilisateur de qualité, nous vous recommandons de minimiser les conflits Excel avec ces bonnes pratiques :

- Utilisez uniquement les raccourcis clavier avec le modèle suivant : **Ctrl+Shift+Alt+* x***, où *x* est une autre touche.
- Si vous avez besoin de raccourcis clavier, consultez la liste des [raccourcis](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)clavier Excel et évitez d’en utiliser dans votre module.
- Lorsque le focus du clavier se trouve à l’intérieur de l’interface utilisateur du module, **Ctrl+Espace et** **Ctrl+Shift+F10** ne fonctionnent pas, car il s’agit de raccourcis d’accessibilité essentiels.
- Sur un ordinateur Windows ou Mac, si la commande « Réinitialiser les préférences de raccourci des macros de Office » n’est pas disponible dans le menu de recherche, l’utilisateur peut ajouter manuellement la commande au ruban en personnalisant le ruban par le biais du menu contexté.

## <a name="customize-the-keyboard-shortcuts-per-platform"></a>Personnaliser les raccourcis clavier par plateforme

Il est possible de personnaliser les raccourcis pour qu’ils soient spécifiques à la plateforme. Voici un exemple de l’objet qui personnalise les raccourcis pour chacune des `shortcuts` plateformes suivantes : `windows` , , `mac` `web` . Notez que vous devez toujours avoir une touche `default` de raccourci pour chaque raccourci.

Dans l’exemple suivant, la clé est la clé de retour pour toute `default` plateforme qui n’est pas spécifiée. La seule plateforme non spécifiée est Windows, donc la clé `default` s’applique uniquement aux Windows.

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

Si votre add-in prend en charge plusieurs paramètres régionaux, vous devez trouver la propriété des `name` objets d’action. En outre, si l’un des paramètres régionaux que le add-in prend en charge a des alphabets ou des systèmes d’écriture différents, et par conséquent différents claviers, vous devrez peut-être également trouver les raccourcis. Pour plus d’informations sur la façon de trouver les raccourcis clavier JSON, voir [Localize extended overrides](../develop/localization.md#localize-extended-overrides).

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>Raccourcis de navigateur qui ne peuvent pas être préférés

Lorsque vous utilisez des raccourcis clavier personnalisés sur le web, certains raccourcis clavier utilisés par le navigateur ne peuvent pas être préférés par les modules. Cette liste est un travail en cours. Si vous découvrez d’autres combinaisons qui ne peuvent pas être overridées, faites-le nous savoir à l’aide de l’outil de commentaires en bas de cette page.

- Ctrl+N
- Ctrl+Shift+N
- Ctrl+T
- Ctrl+Shift+T
- Ctrl+W
- Ctrl+PgUp/PgDn

## <a name="next-steps"></a>Étapes suivantes

- Consultez [l Excel exemple de raccourcis](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) clavier.
- Obtenez une vue d’ensemble de l’utilisation des substitutions étendues dans [Work avec des substitutions étendues du manifeste.](../develop/extended-overrides.md)
