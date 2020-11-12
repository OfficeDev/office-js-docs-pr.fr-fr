---
title: Raccourcis clavier personnalisés dans les compléments Office
description: Découvrez comment ajouter des raccourcis clavier personnalisés, également appelés combinaisons de touches, à votre complément Office.
ms.date: 11/09/2020
localization_priority: Normal
ms.openlocfilehash: f95c26067203a4ec2659aa6a632403c96ed81674
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996688"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a>Ajouter des raccourcis clavier personnalisés à vos compléments Office (aperçu)

Les raccourcis clavier, également appelés combinaisons de touches, permettent aux utilisateurs de votre complément de travailler plus efficacement et améliorent l’accessibilité du complément pour les utilisateurs présentant un handicap en fournissant une alternative à la souris.

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> Pour commencer avec une version de travail d’un complément avec des raccourcis clavier déjà activés, clonez et exécutez l’exemple de [raccourcis clavier Excel](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts). Lorsque vous êtes prêt à ajouter des raccourcis clavier à votre propre complément, poursuivez avec cet article.

Il y a trois étapes pour ajouter des raccourcis clavier à un complément :

1. [Configurez le manifeste du complément](#configure-the-manifest).
1. [Créez ou modifiez le fichier JSON de raccourcis](#create-or-edit-the-shortcuts-json-file) pour définir les actions et leurs raccourcis clavier.
1. [Ajoutez un ou plusieurs appels d’exécution](#create-a-mapping-of-actions-to-their-functions) de l’API [Office. actions. Associate](/javascript/api/office/office.actions#associate) pour mapper une fonction à chaque action.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Il y a deux modifications mineures à effectuer par le manifeste. La première consiste à autoriser le complément à utiliser un runtime partagé et l’autre à pointer vers un fichier au format JSON où vous avez défini les raccourcis clavier.

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurer le complément pour utiliser un runtime partagé

L’ajout de raccourcis clavier personnalisés nécessite que votre complément utilise le runtime partagé. Pour plus d’informations, [configurez un complément de sorte qu’il utilise un runtime partagé](../excel/configure-your-add-in-to-use-a-shared-runtime.md).

### <a name="link-the-mapping-file-to-the-manifest"></a>Lier le fichier de mappage au manifeste

Immédiatement en *dessous* (pas à l’intérieur) `<VersionOverrides>` de l’élément dans le manifeste, ajoutez un élément [ExtendedOverrides](../reference/manifest/extendedoverrides.md) . Définissez l' `Url` attribut sur l’URL complète d’un fichier JSON dans votre projet que vous allez créer dans une étape ultérieure.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>Créer ou modifier le fichier JSON des raccourcis

Créez un fichier JSON dans votre projet. Assurez-vous que le chemin d’accès du fichier correspond à l’emplacement que vous avez spécifié pour l' `Url` attribut de l’élément [ExtendedOverrides](../reference/manifest/extendedoverrides.md) . Ce fichier décrit les raccourcis clavier et les actions qu’ils vont appeler.

1. Dans le fichier JSON, il existe deux tableaux. Le tableau actions contient les objets qui définissent les actions à appeler et le tableau de raccourcis contient les objets qui mappent les combinaisons de clés sur les actions. Voici un exemple :

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

    Pour plus d’informations sur les objets JSON, voir [construction des objets action](#constructing-the-action-objects) et [création des objets Shortcut](#constructing-the-shortcut-objects). Le schéma complet pour les raccourcis JSON se trouve [ surextended-manifest.schema.js](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json).

    > [!NOTE]
    > Vous pouvez utiliser le « contrôle » à la place de « CTRL » tout au long de cet article.

    Dans une étape ultérieure, les actions seront elles-mêmes mappées vers les fonctions que vous écrivez. Dans cet exemple, vous allez mapper ultérieurement SHOWTASKPANE à une fonction qui appelle la `Office.addin.showAsTaskpane` méthode et HIDETASKPANE à une fonction qui appelle la `Office.addin.hide` méthode.

## <a name="create-a-mapping-of-actions-to-their-functions"></a>Créer un mappage des actions sur leurs fonctions

1. Dans votre projet, ouvrez le fichier JavaScript chargé par votre page HTML dans l' `<FunctionFile>` élément.
1. Dans le fichier JavaScript, utilisez l’API [Office. actions. Associate](/javascript/api/office/office.actions#associate) pour mapper chaque action que vous avez spécifiée dans le fichier JSON à une fonction JavaScript. Ajoutez le code JavaScript suivant au fichier. Notez ce qui suit dans le code :

    - Le premier paramètre est l’une des actions du fichier JSON.
    - Le deuxième paramètre est la fonction qui s’exécute lorsqu’un utilisateur appuie sur la combinaison de touches mappée à l’action dans le fichier JSON.

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. Pour continuer l’exemple, utilisez `'SHOWTASKPANE'` comme premier paramètre.
1. Pour le corps de la fonction, utilisez la méthode [Office. AddIn. showTaskpane](/javascript/api/office/office.addin.md#showastaskpane--) pour ouvrir le volet Office du complément. Une fois que vous avez fini, le code doit ressembler à ce qui suit :

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

1. Ajoutez un deuxième appel de `Office.actions.associate` fonction pour mapper l' `HIDETASKPANE` action sur une fonction qui appelle [Office. AddIn. Hide](/javascript/api/office/office.addin.md#hide--). Voici un exemple :

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

Le suivi des étapes précédentes permet à votre complément de faire basculer la visibilité du volet des tâches en appuyant sur **Ctrl + Maj + Flèche vers le haut** et **Ctrl + Maj + Flèche vers le bas**. Il s’agit du même comportement que celui présenté dans l' [exemple de complément Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).

## <a name="details-and-restrictions"></a>Détails et restrictions

### <a name="constructing-the-action-objects"></a>Construction des objets action

Utilisez les instructions suivantes lorsque vous spécifiez les objets dans le `action` tableau de la shortcuts.jssur :

- Les noms `id` de propriété et `name` sont obligatoires.
- La `id` propriété sert à identifier de manière unique l’action à appeler à l’aide d’un raccourci clavier.
- La `name` propriété doit être une chaîne conviviale décrivant l’action. Il doit s’agir d’une combinaison des caractères A-Z, a-z, 0-9 et des signes de ponctuation « - », « _ » et « + ».
- La propriété `type` est facultative. Le type actuellement uniquement `ExecuteFunction` est pris en charge.

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

Le schéma complet pour les raccourcis JSON se trouve [ surextended-manifest.schema.js](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json).

### <a name="constructing-the-shortcut-objects"></a>Construction des objets Shortcut

Utilisez les instructions suivantes lorsque vous spécifiez les objets dans le `shortcuts` tableau de la shortcuts.jssur :

- Les noms des propriétés `action` , `key` et `default` sont obligatoires.
- La valeur de la `action` propriété est une chaîne qui doit correspondre à l’une des `id` Propriétés de l’objet action.
- La `default` propriété peut être n’importe quelle combinaison des caractères a-z, a-z, 0-9 et des signes de ponctuation « - », « _ » et « + ». (Par Convention, les lettres minuscules ne sont pas utilisées dans ces propriétés.)
- La `default` propriété doit contenir le nom d’au moins une touche de modification (Alt, CTRL, Maj) et une seule autre clé.
- Pour Mac, nous pouvons également prendre en charge la touche de modification de commande.
- Pour Mac, la touche ALT est mappée sur la touche d’OPTION. Pour Windows, la commande est mappée sur la touche CTRL.
- Lorsque deux caractères sont liés à la même clé physique dans un clavier standard, alors qu’il s’agit de synonymes dans la `default` propriété ; par exemple, Alt + a et Alt + a représentent le même raccourci, donc Ctrl +-et Ctrl + \_ car « - » et « _ » sont la même clé physique.
- Le caractère « + » indique que les touches de part et d’autre de celle-ci sont enfoncées simultanément.

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

Le schéma complet pour les raccourcis JSON se trouve [ surextended-manifest.schema.js](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json).

> [!NOTE]
> Les KeyTips, également appelés raccourcis clavier séquentiels, tels que le raccourci Excel pour choisir une couleur de remplissage **Alt + h, h** , ne sont pas pris en charge dans les compléments Office.

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a>Utilisation de raccourcis lorsque le volet Office est sélectionné

Actuellement, les raccourcis clavier d’un complément Office ne peuvent être appelés que lorsque le focus de l’utilisateur se trouve dans la feuille de calcul. Lorsque le focus de l’utilisateur se trouve à l’intérieur de l’interface utilisateur Office (par exemple, le volet Office), aucun des raccourcis du complément n’est ignoré. En guise de solution de contournement, le complément peut définir des gestionnaires de clavier qui peuvent appeler certaines actions lorsque l’utilisateur se trouve à l’intérieur de l’interface utilisateur du complément.

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a>Utilisation de combinaisons de touches déjà utilisées par Office ou un autre complément

Pendant la période de préversion, il n’existe aucun système permettant de déterminer ce qui se produit lorsqu’un utilisateur appuie sur une combinaison de touches enregistrée par un complément et par Office ou par un autre complément. Le comportement n’est pas défini.

Actuellement, il n’existe aucune solution de contournement lorsque deux ou plusieurs compléments ont inscrit le même raccourci clavier, mais vous pouvez réduire les conflits avec Excel avec ces bonnes pratiques :

- Utilisez uniquement des raccourcis clavier avec le modèle suivant dans votre complément : * *Ctrl + Maj + Alt +* x * * *, où *x* est une autre touche.
- Si vous avez besoin de raccourcis clavier supplémentaires, consultez la [liste des raccourcis clavier Excel](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)et évitez de les utiliser dans votre complément.

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>Raccourcis de navigateur qui ne peuvent pas être remplacés

Vous ne pouvez pas utiliser l’une des combinaisons de touches suivantes. Elles sont utilisées par les navigateurs et ne peuvent pas être remplacées. Cette liste est un travail en cours. Si vous découvrez d’autres combinaisons qui ne peuvent pas être remplacées, faites-le nous savoir en utilisant l’outil de commentaires en bas de cette page.

- Ctrl + N
- Ctrl + Maj + N
- Ctrl + T
- Ctrl + Maj + T
- CTRL + W
- Ctrl + PG. préc/PG. suiv

## <a name="next-steps"></a>Étapes suivantes

- Voir l’exemple de complément [Excel-clavier-raccourcis](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).
