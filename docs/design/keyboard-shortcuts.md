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
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a>Ajouter des raccourcis clavier personnalisés à vos add-ins Office (aperçu)

Les raccourcis clavier, également appelés combinaisons de touches, permettent aux utilisateurs de votre module de travailler plus efficacement et améliorent l’accessibilité du module pour les utilisateurs présentant un handicap en offrant une alternative à la souris.

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> Pour commencer avec une version de travail d’un add-in avec des raccourcis clavier déjà activés, clonez et exécutez l’exemple de [raccourcis clavier Excel.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) Lorsque vous êtes prêt à ajouter des raccourcis clavier à votre propre add-in, poursuivez avec cet article.

Il existe trois étapes pour ajouter des raccourcis clavier à un add-in :

1. [Configurez le manifeste du add-in.](#configure-the-manifest)
1. [Créez ou modifiez le fichier JSON](#create-or-edit-the-shortcuts-json-file) de raccourcis pour définir des actions et leurs raccourcis clavier.
1. [Ajoutez un ou plusieurs appels d’runtime](#create-a-mapping-of-actions-to-their-functions) de l’API [Office.actions.associate](/javascript/api/office/office.actions#associate) pour ma cartographier une fonction à chaque action.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Deux petites modifications sont à apporter au manifeste. L’une consiste à permettre au add-in d’utiliser un runtime partagé et l’autre à pointer vers un fichier au format JSON où vous avez défini les raccourcis clavier.

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurer le add-in pour utiliser un runtime partagé

L’ajout de raccourcis clavier personnalisés nécessite que votre add-in utilise le runtime partagé. Pour plus d’informations, [configurez un module complémentaire pour utiliser un runtime partagé.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)

### <a name="link-the-mapping-file-to-the-manifest"></a>Lier le fichier de mappage au manifeste

Juste *en dessous* (pas à l’intérieur) de l’élément dans le manifeste, ajoutez un élément `<VersionOverrides>` [ExtendedOverrides.](../reference/manifest/extendedoverrides.md) Définissez l’attribut sur l’URL complète d’un fichier JSON dans votre projet que vous `Url` créerez à une étape ultérieure.

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

    Pour plus d’informations sur les objets JSON, voir [Constructing the action objects](#constructing-the-action-objects) and [Constructing the shortcut objects](#constructing-the-shortcut-objects). The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

    > [!NOTE]
    > Vous pouvez utiliser « CONTROL » à la place de « Ctrl » tout au long de cet article.

    Dans une étape ultérieure, les actions seront elles-mêmes mappées aux fonctions que vous écrivez. Dans cet exemple, vous masquez ultérieurement SHOWTASKPANE à une fonction qui appelle la méthode et HIDETASKPANE à une fonction qui `Office.addin.showAsTaskpane` appelle la `Office.addin.hide` méthode.

## <a name="create-a-mapping-of-actions-to-their-functions"></a>Créer un mappage des actions à leurs fonctions

1. Dans votre projet, ouvrez le fichier JavaScript chargé par votre page HTML dans `<FunctionFile>` l’élément.
1. Dans le fichier JavaScript, utilisez l’API [Office.actions.associate](/javascript/api/office/office.actions#associate) pour ma cartographier chaque action que vous avez spécifiée dans le fichier JSON sur une fonction JavaScript. Ajoutez le javaScript suivant au fichier. Notez ce qui suit à propos du code :

    - Le premier paramètre est l’une des actions du fichier JSON.
    - Le deuxième paramètre est la fonction qui s’exécute lorsqu’un utilisateur appuie sur la combinaison de touches mappée à l’action dans le fichier JSON.

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. Pour continuer l’exemple, `'SHOWTASKPANE'` utilisez-le comme premier paramètre.
1. Pour le corps de la fonction, utilisez la méthode [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) pour ouvrir le volet Office du add-in. Lorsque vous avez terminé, le code doit ressembler à ce qui suit :

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

1. Ajoutez un deuxième appel de `Office.actions.associate` fonction pour maque `HIDETASKPANE` l’action sur une fonction qui appelle [Office.addin.hide](/javascript/api/office/office.addin#hide--). Voici un exemple :

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

La suite des étapes précédentes permet à votre add-in de faire tourner la visibilité du volet Des tâches en appuyant sur **Ctrl+Shift+Flèche** vers le haut et **Ctrl+Shift+Flèche** vers le bas. Il s’agit du même comportement que dans l’exemple de [raccourcis clavier Excel.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)

## <a name="details-and-restrictions"></a>Détails et restrictions

### <a name="constructing-the-action-objects"></a>Construction des objets d’action

Utilisez les instructions suivantes lors de la spécification des objets dans le tableau de la `action` shortcuts.jssuivantes :

- Les noms des `id` propriétés `name` sont obligatoires.
- La `id` propriété est utilisée pour identifier de manière unique l’action à appeler à l’aide d’un raccourci clavier.
- La `name` propriété doit être une chaîne conviviale décrivant l’action. Il doit s’agit d’une combinaison des caractères A - Z, a - z, 0 - 9 et des signes de ponctuation « - », « _ » et « + ».
- La propriété `type` est facultative. Actuellement, `ExecuteFunction` seul le type est pris en charge.

Voici un exemple :

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

### <a name="constructing-the-shortcut-objects"></a>Construction des objets de raccourci

Utilisez les instructions suivantes lors de la spécification des objets dans le tableau de la `shortcuts` shortcuts.jssuivantes :

- Les noms des `action` propriétés `key` et sont `default` obligatoires.
- La valeur de la propriété est une chaîne et doit correspondre à l’une `action` des `id` propriétés de l’objet action.
- La propriété peut être n’importe quelle combinaison des caractères `default` A - Z, -z, 0 - 9 et les signes de ponctuation « - », « _ » et « + ». (Par convention, les lettres majuscules ne sont pas utilisées dans ces propriétés.)
- La propriété doit contenir le nom d’au moins une touche de `default` modification (ALT, Ctrl, SHIFT) et une seule autre touche.
- Pour les Mac, nous prise en charge également la touche modificateur COMMAND.
- Pour les Mac, Alt est mappée sur la touche OPTION. Pour Windows, la commande est mappée sur la touche Ctrl.
- Lorsque deux caractères sont liés à la même touche physique dans un clavier standard, ils sont synonymes dans la propriété ; par exemple, ALT+a et ALT+A sont le même raccourci, tout comme `default` Ctrl+- et Ctrl+ car « - » et « _ » sont la même touche \_ physique.
- Le caractère « + » indique que les touches de chaque côté de celui-ci sont entrées simultanément.

Voici un exemple :

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

The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

> [!NOTE]
> Les touches d’accès, également appelées raccourcis clavier séquentiels, tels que le raccourci Excel pour choisir une couleur de remplissage **Alt+H, H,** ne sont pas pris en charge dans les add-ins Office.

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a>Utilisation de raccourcis lorsque le focus se trouve dans le volet Des tâches

Actuellement, les raccourcis clavier d’un add-in Office ne peuvent être appelés que lorsque le focus de l’utilisateur se trouve dans la feuille de calcul. Lorsque le focus de l’utilisateur se trouve à l’intérieur de l’interface utilisateur d’Office (par exemple, le volet Office), aucun des raccourcis du add-in n’est ignoré. Comme solution de contournement, le add-in peut définir des handlers de clavier qui peuvent appeler certaines actions lorsque le focus de l’utilisateur est à l’intérieur de l’interface utilisateur du add-in.

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a>Utilisation de combinaisons de touches déjà utilisées par Office ou un autre module

Pendant la période d’aperçu, il n’existe aucun système permettant de déterminer ce qui se produit lorsqu’un utilisateur appuie sur une combinaison de touches inscrite par un module et par Office ou par un autre. Le comportement n’est pas définie.

Actuellement, il n’existe aucune solution de contournement lorsque deux ou plusieurs modules ont inscrit le même raccourci clavier, mais vous pouvez réduire les conflits avec Excel avec ces bonnes pratiques :

- Utilisez uniquement les raccourcis clavier avec le modèle suivant dans votre add-in : **Ctrl+Shift+Alt+* x***, où *x* est une autre touche.
- Si vous avez besoin de raccourcis clavier, consultez la liste des [raccourcis](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)clavier Excel et évitez d’en utiliser un dans votre module.

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>Raccourcis du navigateur qui ne peuvent pas être préférés

Vous ne pouvez utiliser aucune des combinaisons de clavier suivantes. Ils sont utilisés par les navigateurs et ne peuvent pas être utilisés. Cette liste est un travail en cours. Si vous découvrez d’autres combinaisons qui ne peuvent pas être overridées, faites-le nous savoir à l’aide de l’outil de commentaires en bas de cette page.

- Ctrl+N
- Ctrl+Shift+N
- Ctrl+T
- Ctrl+Shift+T
- Ctrl+W
- Ctrl+PgUp/PgDn

## <a name="localize-the-keyboard-shortcuts-json"></a>Localisez les raccourcis clavier JSON

Si votre add-in prend en charge plusieurs paramètres régionaux, vous devez trouver la propriété des `name` objets d’action. En outre, si l’un des paramètres régionaux que le add-in prend en charge a des alphabets ou des systèmes d’écriture différents, et par conséquent différents claviers, vous devrez peut-être également trouver les raccourcis. Pour plus d’informations sur la façon de trouver les raccourcis clavier JSON, voir [Localize extended overrides](../develop/localization.md#localize-extended-overrides).

## <a name="next-steps"></a>Étapes suivantes

- Consultez l’exemple de raccourcis [clavier-excel pour le add-in.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)
- Obtenez une vue d’ensemble de l’utilisation des substitutions étendues dans [Work avec des substitutions étendues du manifeste.](../develop/extended-overrides.md)
