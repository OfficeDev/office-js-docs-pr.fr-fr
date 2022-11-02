---
title: Créer des onglets contextuels personnalisés dans les compléments Office
description: Découvrez comment ajouter des onglets contextuels personnalisés à votre complément Office.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1f43f6ec0a6ef3faef4c5e50d5da6d124124fe92
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810231"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a>Créer des onglets contextuels personnalisés dans les compléments Office

Un onglet contextuel est un contrôle d’onglet masqué dans le ruban Office qui s’affiche dans la ligne d’onglet lorsqu’un événement spécifié se produit dans le document Office. Par exemple, l’onglet **Création** de tableau qui apparaît dans le ruban Excel lorsqu’un tableau est sélectionné. Vous incluez des onglets contextuels personnalisés dans votre complément Office et spécifiez quand ils sont visibles ou masqués, en créant des gestionnaires d’événements qui modifient la visibilité. (Toutefois, les onglets contextuels personnalisés ne répondent pas aux modifications de focus.)

> [!NOTE]
> Cet article suppose que vous connaissez la documentation décrite ci-après. Étudiez-la si vous n’avez pas récemment utilisé les commandes de complément (éléments de menu et boutons de ruban personnalisés).
>
> - [Concepts basiques pour les commandes de complément](add-in-commands.md)

> [!IMPORTANT]
> Les onglets contextuels personnalisés sont actuellement pris en charge uniquement sur Excel et uniquement sur ces plateformes et builds.
>
> - Excel sur Windows : version 2102 (build 13801.20294) ou ultérieure.
> - Excel sur Mac : version 16.53.806.0 ou ultérieure.
> - Excel sur le web

> [!NOTE]
> Les onglets contextuels personnalisés fonctionnent uniquement sur les plateformes qui prennent en charge les ensembles de conditions requises suivants. Pour plus d’informations sur les ensembles de conditions requises et leur utilisation, consultez [Spécifier les applications Office et les exigences d’API](../develop/specify-office-hosts-and-api-requirements.md).
>
> - [RibbonApi 1.2](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets)
> - [SharedRuntime 1.1](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)
>
> Vous pouvez utiliser les vérifications d’exécution dans votre code pour tester si la combinaison de l’hôte et de la plateforme de l’utilisateur prend en charge ces ensembles de conditions requises, comme décrit dans [Vérifications au moment de l’exécution pour la prise en charge de la méthode et de l’ensemble de conditions requises](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). (La technique de spécification des ensembles de conditions requises dans le manifeste, qui est également décrite dans cet article, ne fonctionne pas actuellement pour RibbonApi 1.2.) Vous pouvez également [implémenter une autre expérience d’interface utilisateur lorsque les onglets contextuels personnalisés ne sont pas pris en charge](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

## <a name="behavior-of-custom-contextual-tabs"></a>Comportement des onglets contextuels personnalisés

L’expérience utilisateur pour les onglets contextuels personnalisés suit le modèle des onglets contextuels Office intégrés. Voici les principes de base pour les onglets contextuels personnalisés de placement.

- Lorsqu’un onglet contextuel personnalisé est visible, il apparaît à l’extrémité droite du ruban.
- Si un ou plusieurs onglets contextuels intégrés et un ou plusieurs onglets contextuels personnalisés des compléments sont visibles en même temps, les onglets contextuels personnalisés sont toujours à droite de tous les onglets contextuels intégrés.
- Si votre complément comporte plusieurs onglets contextuels et que plusieurs contextes sont visibles, ils apparaissent dans l’ordre dans lequel ils sont définis dans votre complément. (La direction est la même que la langue d’Office; autrement dit, est de gauche à droite dans les langues de gauche à droite, mais de droite à gauche dans les langues de droite à gauche.) Pour plus [d’informations sur la façon dont vous les définissez, consultez Définir les groupes et les contrôles qui s’affichent sous l’onglet](#define-the-groups-and-controls-that-appear-on-the-tab) .
- Si plusieurs compléments ont un onglet contextuel visible dans un contexte spécifique, ils apparaissent dans l’ordre dans lequel les compléments ont été lancés.
- Les onglets *contextuels* personnalisés, contrairement aux onglets principaux personnalisés, ne sont pas ajoutés définitivement au ruban de l’application Office. Ils sont présents uniquement dans les documents Office sur lesquels votre complément est en cours d’exécution.

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>Étapes principales pour inclure un onglet contextuel dans un complément

Voici les principales étapes pour inclure un onglet contextuel personnalisé dans un complément.

1. Configurez le complément pour utiliser un runtime partagé.
1. Définissez l’onglet et les groupes et contrôles qui s’y affichent.
1. Inscrivez l’onglet contextuel auprès d’Office.
1. Spécifiez les circonstances dans lesquelles l’onglet sera visible.

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurer le complément pour utiliser un runtime partagé

L’ajout d’onglets contextuels personnalisés nécessite que votre complément utilise le [runtime partagé](../testing/runtimes.md#shared-runtime). Pour plus d’informations, consultez [Configurer un complément pour utiliser un runtime partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>Définir les groupes et les contrôles qui s’affichent sous l’onglet

Contrairement aux onglets principaux personnalisés, qui sont définis avec du code XML dans le manifeste, les onglets contextuels personnalisés sont définis au moment de l’exécution avec un objet blob JSON. Votre code analyse l’objet blob dans un objet JavaScript, puis passe l’objet à la méthode [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)) . Les onglets contextuels personnalisés sont uniquement présents dans les documents sur lesquels votre complément est en cours d’exécution. Cela diffère des onglets principaux personnalisés qui sont ajoutés au ruban de l’application Office lors de l’installation du complément et restent présents lors de l’ouverture d’un autre document. En outre, la `requestCreateControls` méthode ne peut être exécutée qu’une seule fois dans une session de votre complément. Si elle est appelée à nouveau, une erreur est générée.

> [!NOTE]
> La structure des propriétés et sous-propriétés de l’objet blob JSON (ainsi que les noms de clés) est à peu près parallèle à la structure de l’élément [CustomTab](/javascript/api/manifest/customtab) et de ses éléments descendants dans le manifeste XML.

Nous allons construire un exemple d’objet blob JSON d’onglets contextuels pas à pas. Le schéma complet de l’onglet contextuel JSON se trouve dans [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json). Si vous travaillez dans Visual Studio Code, vous pouvez utiliser ce fichier pour obtenir IntelliSense et valider votre json. Pour plus d’informations, consultez [Modification de JSON avec Visual Studio Code - Schémas et paramètres JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).

1. Commencez par créer une chaîne JSON avec deux propriétés de tableau nommées `actions` et `tabs`. Le `actions` tableau est une spécification de toutes les fonctions qui peuvent être exécutées par des contrôles sous l’onglet contextuel. Le `tabs` tableau définit un ou plusieurs onglets contextuels, *jusqu’à un maximum de 20*.

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. Cet exemple simple d’un onglet contextuel n’aura qu’un seul bouton et, par conséquent, une seule action. Ajoutez ce qui suit en tant que seul membre du `actions` tableau. À propos de ce balisage, notez :

    - Les `id` propriétés et `type` sont obligatoires.
    - La valeur de `type` peut être « ExecuteFunction » ou « ShowTaskpane ».
    - La `functionName` propriété est utilisée uniquement lorsque la valeur de `type` est `ExecuteFunction`. Il s’agit du nom d’une fonction définie dans le FunctionFile. Pour plus d’informations sur functionFile, consultez [Concepts de base pour les commandes de complément](add-in-commands.md).
    - Dans une étape ultérieure, vous mapperez cette action à un bouton sous l’onglet contextuel.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
    ```

1. Ajoutez ce qui suit en tant que seul membre du `tabs` tableau. À propos de ce balisage, notez :

    - La propriété `id` est requise. Utilisez un ID bref et descriptif unique parmi tous les onglets contextuels de votre complément.
    - La propriété `label` est requise. Il s’agit d’une chaîne conviviale qui sert d’étiquette de l’onglet contextuel.
    - La propriété `groups` est requise. Il définit les groupes de contrôles qui s’afficheront sous l’onglet . Il doit avoir au moins un membre *et pas plus de 20*. (Il existe également des limites sur le nombre de contrôles que vous pouvez avoir sur un onglet contextuel personnalisé, ce qui limite également le nombre de groupes dont vous disposez. Pour plus d’informations, consultez l’étape suivante.)

    > [!NOTE]
    > L’objet tab peut également avoir une propriété facultative `visible` qui spécifie si l’onglet est visible immédiatement au démarrage du complément. Étant donné que les onglets contextuels sont normalement masqués jusqu’à `false` ce qu’un événement utilisateur déclenche leur visibilité (par exemple, l’utilisateur sélectionnant une entité d’un type quelconque dans le document), la propriété prend par défaut la `visible` valeur lorsqu’elle n’est pas présente. Dans une section ultérieure, nous montrons comment définir la propriété sur `true` en réponse à un événement.

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. Dans l’exemple simple en cours, l’onglet contextuel n’a qu’un seul groupe. Ajoutez ce qui suit en tant que seul membre du `groups` tableau. À propos de ce balisage, notez :

    - Toutes les propriétés sont requises.
    - La `id` propriété doit être unique parmi tous les groupes du manifeste. Utilisez un ID bref et descriptif de 125 caractères maximum.
    - est `label` une chaîne conviviale qui sert d’étiquette du groupe.
    - La `icon` valeur de la propriété est un tableau d’objets qui spécifient les icônes que le groupe aura sur le ruban en fonction de la taille du ruban et de la fenêtre de l’application Office.
    - La `controls` valeur de la propriété est un tableau d’objets qui spécifient les boutons et les menus du groupe. Il doit y en avoir au moins un.

    > [!IMPORTANT]
    > *Le nombre total de contrôles sur l’ensemble de l’onglet ne peut pas être supérieur à 20.* Par exemple, vous pouvez avoir 3 groupes avec 6 contrôles chacun et un quatrième groupe avec 2 contrôles, mais vous ne pouvez pas avoir 4 groupes avec 6 contrôles chacun.  

    ```json
    {
        "id": "CustomGroup111",
        "label": "Insertion",
        "icon": [

        ],
        "controls": [

        ]
    }
    ```

1. Chaque groupe doit avoir une icône d’au moins deux tailles, 32 x 32 px et 80 x 80 px. Si vous le souhaitez, vous pouvez également avoir des icônes de tailles 16 x 16 px, 20 x 20 px, 24 x 24 px, 40 x 40 px, 48 x 48 px et 64 x 64 px. Office décide de l’icône à utiliser en fonction de la taille du ruban et de la fenêtre de l’application Office. Ajoutez les objets suivants au tableau d’icônes. (Si les tailles de fenêtre et de ruban sont suffisamment grandes pour qu’au moins l’un des *contrôles* du groupe s’affiche, aucune icône de groupe n’apparaît. Pour obtenir un exemple, regardez le groupe **Styles** sur le ruban Word à mesure que vous réduisez et développez la fenêtre Word.) À propos de ce balisage, notez :

    - Les deux propriétés sont obligatoires.
    - L’unité `size` de propriété de mesure est pixels. Les icônes étant toujours carrées, le nombre correspond à la fois à la hauteur et à la largeur.
    - La `sourceLocation` propriété spécifie l’URL complète de l’icône.

    > [!IMPORTANT]
    > Tout comme vous devez généralement modifier les URL dans le manifeste du complément lorsque vous passez du développement à la production (par exemple, en changeant le domaine de localhost à contoso.com), vous devez également modifier les URL dans vos onglets contextuels JSON.

    ```json
    {
        "size": 32,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
    },
    {
        "size": 80,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
    }
    ```

1. Dans notre exemple simple en cours, le groupe n’a qu’un seul bouton. Ajoutez l’objet suivant comme seul membre du `controls` tableau. À propos de ce balisage, notez :

    - Toutes les propriétés, à l’exception `enabled`de , sont obligatoires.
    - `type` spécifie le type de contrôle. Les valeurs peuvent être « Button », « Menu » ou « MobileButton ».
    - `id` peut comporter jusqu’à 125 caractères.
    - `actionId` doit être l’ID d’une action définie dans le `actions` tableau. (Consultez l’étape 1 de cette section.)
    - `label` est une chaîne conviviale qui sert d’étiquette du bouton.
    - `superTip` représente une forme enrichie d’info-bulle. `title` Les propriétés et sont `description` requises.
    - `icon` spécifie les icônes du bouton. Les remarques précédentes sur l’icône de groupe s’appliquent également ici.
    - `enabled` (facultatif) spécifie si le bouton est activé lorsque l’onglet contextuel apparaît au démarrage. La valeur par défaut s’il n’est pas présent est `true`.

    ```json
    {
        "type": "Button",
        "id": "CtxBt112",
        "actionId": "executeWriteData",
        "enabled": false,
        "label": "Write Data",
        "superTip": {
            "title": "Data Insertion",
            "description": "Use this button to insert data into the document."
        },
        "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
            }
        ]
    }
    ```

Voici l’exemple complet de l’objet blob JSON.

```json
`{
  "actions": [
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
  ],
  "tabs": [
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [
        {
          "id": "CustomGroup111",
          "label": "Insertion",
          "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
            }
          ],
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "executeWriteData",
                "enabled": false,
                "label": "Write Data",
                "superTip": {
                    "title": "Data Insertion",
                    "description": "Use this button to insert data into the document."
                },
                "icon": [
                    {
                        "size": 32,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
                    },
                    {
                        "size": 80,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
                    }
                ]
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>Inscrire l’onglet contextuel auprès d’Office avec requestCreateControls

L’onglet contextuel est inscrit auprès d’Office en appelant la méthode [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)) . Cette opération est généralement effectuée dans la fonction affectée à `Office.initialize` ou avec la `Office.onReady` fonction . Pour plus d’informations sur ces fonctions et l’initialisation du complément, voir [Initialiser votre complément Office](../develop/initialize-add-in.md). Toutefois, vous pouvez appeler la méthode à tout moment après l’initialisation.

> [!IMPORTANT]
> La `requestCreateControls` méthode ne peut être appelée qu’une seule fois dans une session donnée d’un complément. Une erreur est générée si elle est appelée à nouveau.

Voici un exemple. Notez que la chaîne JSON doit être convertie en objet JavaScript avec la `JSON.parse` méthode avant de pouvoir être passée à une fonction JavaScript.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>Spécifier les contextes où l’onglet sera visible avec requestUpdate

En règle générale, un onglet contextuel personnalisé doit apparaître lorsqu’un événement initié par l’utilisateur modifie le contexte du complément. Imaginez un scénario dans lequel l’onglet doit être visible lorsque, et uniquement quand, un graphique (sur la feuille de calcul par défaut d’un classeur Excel) est activé.

Commencez par attribuer des gestionnaires. Cela est généralement effectué dans la `Office.onReady` fonction, comme dans l’exemple suivant, qui affecte des gestionnaires (créés à une étape ultérieure) aux `onActivated` événements et `onDeactivated` de tous les graphiques de la feuille de calcul.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

    await Excel.run(context => {
        const charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

Ensuite, définissez les gestionnaires. Voici un exemple simple de , mais consultez Gestion de `showDataTab`[l’erreur HostRestartNeeded plus](#handle-the-hostrestartneeded-error) loin dans cet article pour obtenir une version plus robuste de la fonction. Tenez compte du code suivant :

- Office effectue un contrôle lorsqu’il met à jour l’état du ruban. La méthode  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)) met en file d’attente une demande de mise à jour. La méthode résout l’objet `Promise` dès qu’elle a mis la demande en file d’attente, et non lorsque le ruban est réellement mis à jour.
- Le paramètre de la `requestUpdate` méthode est un objet [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) qui (1) spécifie l’onglet par son ID *exactement comme spécifié dans le JSON* et (2) spécifie la visibilité de l’onglet.
- Si vous avez plusieurs onglets contextuels personnalisés qui doivent être visibles dans le même contexte, vous ajoutez simplement des objets tabulation supplémentaires au `tabs` tableau.

```javascript
async function showDataTab() {
    await Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            }
        ]});
}
```

Le gestionnaire permettant de masquer l’onglet est presque identique, sauf qu’il rétablit la `visible` propriété sur `false`.

La bibliothèque JavaScript Office fournit également plusieurs interfaces (types) pour faciliter la construction de l’objet`RibbonUpdateData` . La fonction suivante dans `showDataTab` TypeScript utilise ces types.

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Activer la visibilité des onglets et l’état activé d’un bouton en même temps

La `requestUpdate` méthode est également utilisée pour activer ou désactiver l’état activé ou désactivé d’un bouton personnalisé sur un onglet contextuel personnalisé ou un onglet principal personnalisé. Pour plus d’informations à ce sujet, consultez [Activer et désactiver les commandes de complément](disable-add-in-commands.md). Il peut y avoir des scénarios dans lesquels vous souhaitez modifier à la fois la visibilité d’un onglet et l’état activé d’un bouton en même temps. Vous effectuez cette opération avec un seul appel de `requestUpdate`. Voici un exemple dans lequel un bouton sur un onglet principal est activé en même temps qu’un onglet contextuel est rendu visible.

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            },
            {
                id: "OfficeAppTab1",
                groups: [
                    {
                        id: "CustomGroup111",
                        controls: [
                            {
                                id: "MyButton",
                                enabled: true
                            }
                        ]
                    }
                ]
            ]}
        ]
    });
}
```

Dans l’exemple suivant, le bouton activé se trouve sous le même onglet contextuel que celui qui est visible.

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                groups: [
                    {
                        id: "CustomGroup111",
                        controls: [
                            {
                                id: "MyButton",
                                enabled: true
                           }
                       ]
                   }
               ]
            }
        ]
    });
}
```

## <a name="open-a-task-pane-from-contextual-tabs"></a>Ouvrir un volet Office à partir d’onglets contextuels

Pour ouvrir votre volet Office à partir d’un bouton sur un onglet contextuel personnalisé, créez une action dans le JSON avec un `type` de `ShowTaskpane`. Définissez ensuite un bouton avec la `actionId` propriété définie sur le `id` de l’action. Cela ouvre le volet Office par défaut spécifié par l’élément **\<Runtime\>** dans votre manifeste.

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

Pour ouvrir un volet Office qui n’est pas le volet Office par défaut, spécifiez une `sourceLocation` propriété dans la définition de l’action. Dans l’exemple suivant, un deuxième volet office est ouvert à partir d’un autre bouton.

> [!IMPORTANT]
>
> - Quand un `sourceLocation` est spécifié pour l’action, le volet Office n’utilise *pas* le runtime partagé. Il s’exécute dans un nouveau runtime distinct.
> - Pas plus d’un volet Office peut utiliser le runtime partagé, donc pas plus d’une action de type `ShowTaskpane` ne peut omettre la `sourceLocation` propriété.

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    },
    {
      "id": "openTablesTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Tables",
      "supportPinning": false
      "sourceLocation": "https://MyDomain.com/myPage.html"
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            },
            {
                "type": "Button",
                "id": "CtxBt113",
                "actionId": "openTablesTaskpane",
                "enabled": false,
                "label": "Open Tables Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="localize-the-json-text"></a>Localiser le texte JSON

L’objet blob JSON qui est passé à `requestCreateControls` n’est pas localisé de la même façon que le balisage du manifeste pour les onglets principaux personnalisés est localisé (ce qui est décrit dans [Contrôler la localisation à partir du manifeste](../develop/localization.md#control-localization-from-the-manifest)). Au lieu de cela, la localisation doit se produire au moment de l’exécution à l’aide d’objets blob JSON distincts pour chaque paramètre régional. Nous vous suggérons d’utiliser une `switch` instruction qui teste la propriété [Office.context.displayLanguage](/javascript/api/office/office.context#office-office-context-displaylanguage-member) . Voici un exemple.

```javascript
function GetContextualTabsJsonSupportedLocale () {
    const displayLanguage = Office.context.displayLanguage;

        switch (displayLanguage) {
            case 'en-US':
                return `{
                    "actions": [
                        // actions omitted
                     ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Data",
                          "groups": [
                              // groups omitted
                          ]
                        }
                    ]
                }`;

            case 'fr-FR':
                return `{
                    "actions": [
                        // actions omitted 
                    ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Données",
                          "groups": [
                              // groups omitted
                          ]
                       }
                    ]
               }`;

            // Other cases omitted
       }
}
```

Ensuite, votre code appelle la fonction pour obtenir l’objet blob localisé qui est passé à `requestCreateControls`, comme dans l’exemple suivant.

```javascript
const contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>Meilleures pratiques pour les onglets contextuels personnalisés

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>Implémenter une autre expérience d’interface utilisateur lorsque les onglets contextuels personnalisés ne sont pas pris en charge

Certaines combinaisons de plateforme, d’application Office et de build Office ne prennent pas en charge `requestCreateControls`. Votre complément doit être conçu pour fournir une autre expérience aux utilisateurs qui exécutent le complément sur l’une de ces combinaisons. Les sections suivantes décrivent deux façons de fournir une expérience de secours.

#### <a name="use-noncontextual-tabs-or-controls"></a>Utiliser des onglets ou des contrôles non textuels

Il existe un élément manifeste, [OverriddenByRibbonApi](/javascript/api/manifest/overriddenbyribbonapi), conçu pour créer une expérience de secours dans un complément qui implémente des onglets contextuels personnalisés lorsque le complément s’exécute sur une application ou une plateforme qui ne prend pas en charge les onglets contextuels personnalisés.

La stratégie la plus simple pour utiliser cet élément consiste à définir un onglet principal personnalisé (autrement dit, un onglet personnalisé *non contextuel* ) dans le manifeste qui duplique les personnalisations du ruban des onglets contextuels personnalisés dans votre complément. Mais vous ajoutez `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` en tant que premier élément enfant des éléments [Group](/javascript/api/manifest/group), [Control](/javascript/api/manifest/control) et menu **\<Item\>** en double dans les onglets principaux personnalisés. L’effet de cette opération est le suivant :

- Si le complément s’exécute sur une application et une plateforme qui prennent en charge des onglets contextuels personnalisés, les groupes principaux et les contrôles personnalisés n’apparaissent pas sur le ruban. Au lieu de cela, l’onglet contextuel personnalisé est créé lorsque le complément appelle la `requestCreateControls` méthode .
- Si le complément s’exécute sur une application ou une plateforme qui *ne prend pas* en charge `requestCreateControls`, les éléments apparaissent sous l’onglet cœur personnalisé.

Voici un exemple. Notez que « MyButton » s’affiche sous l’onglet principal personnalisé uniquement lorsque les onglets contextuels personnalisés ne sont pas pris en charge. Toutefois, le groupe parent et l’onglet cœur personnalisé s’affichent, que les onglets contextuels personnalisés soient ou non pris en charge.

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>              
              ...
              <Group ...>
                ...
                <Control ... id="Contoso.MyButton1">
                  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
                  ...
                  <Action ...>
...
</OfficeApp>
```

Pour plus d’exemples, consultez [OverriddenByRibbonApi](/javascript/api/manifest/overriddenbyribbonapi).

Lorsqu’un groupe parent ou un menu est marqué avec `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, il n’est pas visible et tout son balisage enfant est ignoré lorsque les onglets contextuels personnalisés ne sont pas pris en charge. Par conséquent, peu importe si l’un de ces éléments enfants a l’élément **\<OverriddenByRibbonApi\>** ou quelle est sa valeur. Cela implique que si un élément de menu ou un contrôle doit être visible dans tous les contextes, non seulement il ne doit pas être marqué avec `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, mais *son menu et son groupe ancêtres ne doivent pas non plus être marqués de cette façon*.

> [!IMPORTANT]
> Ne marquez pas *tous les* éléments enfants d’un groupe ou d’un menu avec `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`. Cela est inutile si l’élément parent est marqué avec `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` pour les raisons indiquées dans le paragraphe précédent. En outre, si vous laissez le **\<OverriddenByRibbonApi\>** sur le parent (ou si vous le définissez sur `false`), le parent s’affiche, que les onglets contextuels personnalisés soient pris en charge ou non, mais il sera vide lorsqu’ils sont pris en charge. Par conséquent, si tous les éléments enfants ne doivent pas apparaître lorsque les onglets contextuels personnalisés sont pris en charge, marquez le parent avec `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>Utiliser des API qui affichent ou masquent un volet Office dans des contextes spécifiés

En guise d’alternative à **\<OverriddenByRibbonApi\>**, votre complément peut définir un volet Office avec des contrôles d’interface utilisateur qui dupliquent les fonctionnalités des contrôles dans un onglet contextuel personnalisé. Utilisez ensuite les méthodes [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-showastaskpane-member(1)) et [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-hide-member(1)) pour afficher le volet Office lorsque l’onglet contextuel aurait été affiché s’il était pris en charge. Pour plus d’informations sur l’utilisation de ces méthodes, voir [Afficher ou masquer le volet Office de votre complément Office](../develop/show-hide-add-in.md).

### <a name="handle-the-hostrestartneeded-error"></a>Gérer l’erreur HostRestartNeeded

Dans certains scénarios, Office ne peut pas mettre à jour le ruban et renvoie une erreur. Par exemple, si le complément est mis à niveau et que le complément mis à niveau dispose d'un autre groupe de commandes de complément personnalisé, l’application Office doit être fermée et ouverte de nouveau. La méthode `requestUpdate` renvoie l'erreur `HostRestartNeeded` jusqu'à ce que cela soit effectué. Votre code doit gérer cette erreur. Voici un exemple de procédure. Dans ce cas, la méthode `reportError` affiche l’erreur à l’utilisateur.

```javascript
function showDataTab() {
    try {
        Office.ribbon.requestUpdate({
            tabs: [
                {
                    id: "CtxTab1",
                    visible: true
                }
            ]});
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, then close and reopen the Office application.");
        }
    }
}
```

## <a name="resources"></a>Ressources

- [Exemple de code : Créer des onglets contextuels personnalisés sur le ruban](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-contextual-tabs)
- Exemple de démonstration communautaire d’onglets contextuels

> [!VIDEO https://www.youtube.com/embed/9tLfm4boQIo]
