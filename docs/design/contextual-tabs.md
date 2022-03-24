---
title: Créer des onglets contextuels personnalisés dans Office de recherche
description: Découvrez comment ajouter des onglets contextuels personnalisés à votre Office de recherche.
ms.date: 03/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: aa301996d653170d02280efbdb7e94733b5dd924
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63742930"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a>Créer des onglets contextuels personnalisés dans Office de recherche

Un onglet contextuel est un contrôle onglet masqué dans le ruban Office qui est affiché dans la ligne d’onglet lorsqu’un événement spécifié se produit dans le document Office document. Par exemple, **l’onglet Création** de table qui apparaît sur Excel ruban lorsqu’un tableau est sélectionné. Vous incluez des onglets contextuels personnalisés dans votre Office et spécifiez quand ils sont visibles ou masqués en créant des handlers d’événements qui modifient la visibilité. (Toutefois, les onglets contextuels personnalisés ne répondent pas aux changements de focus.)

> [!NOTE]
> Cet article suppose que vous connaissez la documentation décrite ci-après. Étudiez-la si vous n’avez pas récemment utilisé les commandes de complément (éléments de menu et boutons de ruban personnalisés).
>
> - [Concepts basiques pour les commandes de complément](add-in-commands.md)

> [!IMPORTANT]
> Les onglets contextuels personnalisés sont actuellement uniquement pris en charge sur Excel et uniquement sur ces plateformes et builds.
>
> - Excel sur Windows (abonnement Microsoft 365 uniquement) : version 2102 (build 13801.20294) ou ultérieure.
> - Excel mac : version 16.53.806.0 ou ultérieure.
> - Excel sur le web

> [!NOTE]
> Les onglets contextuels personnalisés fonctionnent uniquement sur les plateformes qui supportent les ensembles de conditions requises suivants. Pour plus d’informations sur les ensembles de conditions requises et sur la façon de les utiliser, voir [Spécifier Office applications et les conditions requises des API](../develop/specify-office-hosts-and-api-requirements.md).
>
> - [RibbonApi 1.2](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> Vous pouvez utiliser les vérifications à l’runtime dans votre code pour tester si la combinaison hôte et plateforme de l’utilisateur prend en charge ces ensembles de conditions requises comme décrit dans les vérifications runtime pour la prise en charge des méthodes et des ensembles de [conditions requises](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). (La technique de spécification des ensembles de conditions requises dans le manifeste, également décrite dans cet article, ne fonctionne actuellement pas pour RibbonApi 1.2.) Vous pouvez également implémenter [une autre expérience d’interface utilisateur lorsque les onglets contextuels personnalisés ne sont pas pris en charge](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

## <a name="behavior-of-custom-contextual-tabs"></a>Comportement des onglets contextuels personnalisés

L’expérience utilisateur pour les onglets contextuels personnalisés suit le modèle des onglets Office contextuels intégrés. Voici les principes de base pour l’emplacement des onglets contextuels personnalisés.

- Lorsqu’un onglet contextuel personnalisé est visible, il apparaît à l’extrémité droite du ruban.
- Si un ou plusieurs onglets contextuels intégrés et un ou plusieurs onglets contextuels personnalisés des modules sont visibles en même temps, les onglets contextuels personnalisés sont toujours à droite de tous les onglets contextuels intégrés.
- Si votre add-in possède plusieurs onglets contextuels et qu’il existe des contextes dans lesquels plusieurs onglets sont visibles, ils apparaissent dans l’ordre dans lequel ils sont définis dans votre module. (Le sens est identique à celui de la langue Office ; c’est-à-dire de gauche à droite dans les langues de gauche à droite, mais de droite à gauche dans les langues de droite à gauche.) Pour [plus d’informations sur](#define-the-groups-and-controls-that-appear-on-the-tab) leur définition, voir Définir les groupes et les contrôles qui apparaissent sous l’onglet.
- Si plusieurs d’entre eux ont un onglet contextuel visible dans un contexte spécifique, ils apparaissent dans l’ordre dans lequel les modules ont été lancés.
- Contrairement *aux* onglets principaux personnalisés, les onglets contextuels personnalisés ne sont pas ajoutés Office le ruban de l’application. Ils sont présents uniquement dans Office documents sur lesquels votre module est en cours d’exécution.

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>Étapes principales pour l’ajout d’un onglet contextuel dans un add-in

Voici les principales étapes à suivre pour inclure un onglet contextuel personnalisé dans un add-in.

1. Configurez le add-in pour utiliser un runtime partagé.
1. Définissez l’onglet, ainsi que les groupes et les contrôles qui y apparaissent.
1. Inscrivez l’onglet contextuel avec Office.
1. Spécifiez les circonstances dans le cas où l’onglet sera visible.

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurer le add-in pour utiliser un runtime partagé

L’ajout d’onglets contextuels personnalisés nécessite que votre add-in utilise le runtime partagé. Pour plus d’informations, voir [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>Définir les groupes et les contrôles qui apparaissent sous l’onglet

Contrairement aux onglets principaux personnalisés, qui sont définis avec du XML dans le manifeste, les onglets contextuels personnalisés sont définis lors de l’runtime avec un blob JSON. Votre code parse le blob dans un objet JavaScript, puis passe l’objet à la [méthode Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)). Les onglets contextuels personnalisés sont uniquement présents dans les documents sur lesquels votre module est en cours d’exécution. Cela est différent des onglets principaux personnalisés qui sont ajoutés au ruban de l’application Office lorsque le module est installé et restent présents à l’ouverture d’un autre document. En outre, la `requestCreateControls` méthode ne peut être exécuté qu’une seule fois dans une session de votre add-in. Si elle est appelée à nouveau, une erreur est lancée.

> [!NOTE]
> La structure des propriétés et sous-propriétés de l’objet blob JSON (et les noms clés) est à peu près parallèle à la structure de l’élément [CustomTab](../reference/manifest/customtab.md) et de ses éléments descendants dans le manifeste XML.

Nous allons créer un exemple d’objet blob JSON onglets contextuel pas à pas. Le schéma complet de l’onglet contextuel JSON se trouve [à dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json). Si vous travaillez dans Visual Studio Code, vous pouvez utiliser ce fichier pour obtenir IntelliSense et valider votre JSON. Pour plus d’informations, voir [Modification de JSON Visual Studio Code - Schémas et paramètres JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).

1. Commencez par créer une chaîne JSON avec deux propriétés de tableau nommées `actions` et `tabs`. Le `actions` tableau est une spécification de toutes les fonctions qui peuvent être exécutées par des contrôles sous l’onglet contextuel. Le `tabs` tableau définit un ou plusieurs onglets contextuels, *jusqu’à un maximum de 20*.

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. Cet exemple simple d’onglet contextuel n’aura qu’un seul bouton et, par conséquent, une seule action. Ajoutez ce qui suit en tant que seul membre du `actions` tableau. À propos de ce markup, notez :

    - Les `id` propriétés `type` et les propriétés sont obligatoires.
    - La valeur peut `type` être « ExecuteFunction » ou « ShowTaskpane ».
    - La `functionName` propriété est utilisée uniquement lorsque la valeur est `type` .`ExecuteFunction` Il s’agit du nom d’une fonction définie dans functionFile. Pour plus d’informations sur FunctionFile, voir [Concepts de base des commandes de module complémentaire](add-in-commands.md).
    - Dans une étape ultérieure, vous allez ma cartographier cette action sur un bouton de l’onglet contextuel.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. Ajoutez ce qui suit en tant que seul membre du `tabs` tableau. À propos de ce markup, notez :

    - La propriété `id` est requise. Utilisez un bref ID descriptif unique parmi tous les onglets contextuels de votre module.
    - La propriété `label` est requise. Il s’agit d’une chaîne conviviale qui sert d’étiquette à l’onglet contextuel.
    - La propriété `groups` est requise. Il définit les groupes de contrôles qui apparaîtront sous l’onglet. Elle doit avoir au moins un *membre et pas plus de 20*. (Il existe également des limites au nombre de contrôles que vous pouvez avoir sur un onglet contextuel personnalisé et qui limitent également le nombre de groupes que vous avez. Pour plus d’informations, voir l’étape suivante.)

    > [!NOTE]
    > L’objet tabulation peut également avoir une propriété facultative `visible` qui spécifie si l’onglet est visible immédiatement au démarrage du module. Dans la mesure où les onglets contextuels sont normalement masqués jusqu’à ce qu’un événement utilisateur déclenche leur visibilité (par exemple, l’utilisateur sélectionnant une entité d’un type dans le document), `visible` `false` la propriété se présente par défaut lorsqu’elle n’est pas présente. Dans une section ultérieure, nous montrons comment définir la `true` propriété en réponse à un événement.

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. Dans l’exemple continu simple, l’onglet contextuel ne possède qu’un seul groupe. Ajoutez ce qui suit en tant que seul membre du `groups` tableau. À propos de ce markup, notez :

    - Toutes les propriétés sont requises.
    - La `id` propriété doit être unique parmi tous les groupes dans le manifeste. Utilisez un bref ID descriptif de 125 caractères au plus.
    - Il `label` s’agit d’une chaîne conviviale qui sert d’étiquette au groupe.
    - La `icon` valeur de la propriété est un tableau d’objets qui spécifient les icônes que le groupe aura sur le ruban en fonction de la taille du ruban et de la fenêtre d’application Office.
    - La `controls` valeur de la propriété est un tableau d’objets qui spécifient les boutons et les menus du groupe. Il doit y en avoir au moins un.

    > [!IMPORTANT]
    > *Le nombre total de contrôles sous l’onglet entier ne peut pas être supérieur à 20.* Par exemple, vous pouvez avoir 3 groupes avec 6 contrôles chacun et un quatrième groupe avec 2 contrôles, mais vous ne pouvez pas avoir 4 groupes avec 6 contrôles chacun.  

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

1. Chaque groupe doit avoir une icône d’au moins deux tailles, 32 x 32 px et 80 x 80 px. Si vous le souhaitez, vous pouvez également avoir des icônes de tailles 16 x 16 px, 20 x 20 px, 24 x 24 px, 40 x 40 px, 48 x 48 px et 64 x 64 px. Office l’icône à utiliser en fonction de la taille du ruban et de la Office’application. Ajoutez les objets suivants au tableau d’icônes. (Si les tailles de la fenêtre et du ruban sont suffisamment grandes pour qu’au moins l’un des contrôles du groupe apparaisse, aucune icône de groupe ne s’affiche. Pour obtenir un exemple, regardez le groupe **Styles** sur le ruban Word lorsque vous réduirez et développez la fenêtre Word.) À propos de ce markup, notez :

    - Les deux propriétés sont obligatoires.
    - L’unité `size` de mesure de propriété est pixels. Les icônes sont toujours carrées, de sorte que le nombre est à la fois la hauteur et la largeur.
    - La `sourceLocation` propriété spécifie l’URL complète de l’icône.

    > [!IMPORTANT]
    > Tout comme vous devez généralement modifier les URL dans le manifeste du add-in lorsque vous passez du développement à la production (par exemple, en modifiant le domaine localhost en contoso.com), vous devez également modifier les URL dans vos onglets contextuels JSON.

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

1. Dans notre exemple simple en cours, le groupe ne possède qu’un seul bouton. Ajoutez l’objet suivant comme seul membre du `controls` tableau. À propos de ce markup, notez :

    - Toutes les propriétés, à l’exception `enabled`de , sont obligatoires.
    - `type` spécifie le type de contrôle. Les valeurs peuvent être « Button », « Menu » ou « MobileButton ».
    - `id` peut prendre jusqu’à 125 caractères.
    - `actionId` doit être l’ID d’une action définie dans le `actions` tableau. (Voir l’étape 1 de cette section.)
    - `label` est une chaîne conviviale qui sert d’étiquette au bouton.
    - `superTip` représente une forme enrichie d’info-conseil. Les propriétés `title` et les propriétés `description` sont requises.
    - `icon` spécifie les icônes du bouton. Les remarques précédentes sur l’icône de groupe s’appliquent également ici.
    - `enabled` (facultatif) indique si le bouton est activé au démarrage de l’onglet contextuel. La valeur par défaut, si elle n’est pas présente, est `true`.

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

Voici l’exemple complet du blob JSON.

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>Inscrire l’onglet contextuel Office avec requestCreateControls

L’onglet contextuel est inscrit auprès Office en appelant [la méthode Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)). Cette tâche est généralement effectuée dans la fonction affectée `Office.initialize` à la méthode ou avec celle-ci `Office.onReady` . Pour plus d’informations sur ces méthodes et sur l’initialisation du Office, voir [Initialize your Office Add-in](../develop/initialize-add-in.md). Vous pouvez toutefois appeler la méthode à tout moment après l’initialisation.

> [!IMPORTANT]
> La `requestCreateControls` méthode ne peut être appelée qu’une seule fois dans une session donnée d’un add-in. Une erreur est lancée si elle est appelée à nouveau.

Voici un exemple. Notez que la chaîne JSON doit être convertie en objet JavaScript `JSON.parse` avec la méthode pour pouvoir être transmise à une fonction JavaScript.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>Spécifier les contextes où l’onglet sera visible avec requestUpdate

En règle générale, un onglet contextuel personnalisé doit apparaître lorsqu’un événement initié par l’utilisateur modifie le contexte du add-in. Envisagez un scénario dans lequel l’onglet doit être visible lorsque, et uniquement quand, un graphique (dans la feuille de calcul par défaut d’un Excel)) est activé.

Commencez par affecter des handlers. Cela est généralement `Office.onReady` effectué dans la méthode comme dans l’exemple suivant qui affecte des handlers (créés à une étape ultérieure) `onActivated` `onDeactivated` aux événements et aux graphiques de la feuille de calcul.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

Ensuite, définissez les handlers. Voici un exemple simple `showDataTab`d’une erreur [HostRestartNeeded](#handle-the-hostrestartneeded-error) plus loin dans cet article pour obtenir une version plus robuste de la fonction. Tenez compte du code suivant :

- Office effectue un contrôle lorsqu’il met à jour l’état du ruban. La [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)) met en file d’attente une demande de mise à jour. La méthode résout l’objet `Promise` dès qu’il a mis la demande en file d’attente, et non lorsque le ruban est réellement mis à jour.
- `requestUpdate` Le paramètre de la méthode est un objet [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) qui (1) spécifie l’onglet par son ID exactement comme spécifié dans le *JSON* et (2) spécifie la visibilité de l’onglet.
- Si vous avez plusieurs onglets contextuels personnalisés qui doivent être visibles dans le même contexte, il vous suffit d’ajouter des objets onglet supplémentaires au `tabs` tableau.

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

Le handler pour masquer l’onglet est presque identique, sauf qu’il définit à `visible` nouveau la propriété sur `false`.

La Office JavaScript fournit également plusieurs interfaces (types) pour faciliter la construction de l’objet`RibbonUpdateData`. Voici la fonction dans `showDataTab` TypeScript qui utilise ces types.

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Activer la visibilité de l’onglet et l’état activé d’un bouton en même temps

La `requestUpdate` méthode est également utilisée pour activer ou désactiver l’état d’un bouton personnalisé sur un onglet contextuel personnalisé ou un onglet principal personnalisé. Pour plus d’informations à ce sujet, voir [Enable and Disable Add-in Commands](disable-add-in-commands.md). Il peut y avoir des scénarios dans lesquels vous souhaitez modifier la visibilité d’un onglet et l’état activé d’un bouton en même temps. Vous le faites avec un seul appel de `requestUpdate`. Voici un exemple dans lequel un bouton d’un onglet principal est activé en même temps qu’un onglet contextuel est rendu visible.

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

Dans l’exemple suivant, le bouton activé se trouve sur le même onglet contextuel que celui qui est rendu visible.

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

## <a name="open-a-task-pane-from-contextual-tabs"></a>Ouvrir un volet Des tâches à partir d’onglets contextuels

Pour ouvrir votre volet Des tâches à partir d’un bouton d’un onglet contextuel personnalisé, créez une action dans le JSON avec une `type` des touches .`ShowTaskpane` Définissez ensuite un bouton dont la `actionId` propriété est définie sur la `id` valeur de l’action. Cela ouvre le volet Des tâches par défaut spécifié par l’élément **Runtime** dans votre manifeste.

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

Pour ouvrir un volet De tâches qui n’est pas le volet Des tâches par défaut, `sourceLocation` spécifiez une propriété dans la définition de l’action. Dans l’exemple suivant, un deuxième volet Des tâches est ouvert à partir d’un autre bouton.

> [!IMPORTANT]
>
> - Lorsqu’une `sourceLocation` valeur est spécifiée pour l’action, le volet Des tâches *n’utilise* pas le runtime partagé. Il s’exécute dans un nouveau runtime JavaScript.
> - Un seul volet De tâches ne peut pas utiliser le runtime partagé, de sorte qu’une seule action de type `ShowTaskpane` ne peut pas omettre la `sourceLocation` propriété.

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

Le blob `requestCreateControls` JSON transmis n’est pas localisée de la même façon que le marques de manifeste pour les onglets principaux personnalisés (qui est décrit dans la localisation du contrôle à partir du [manifeste](../develop/localization.md#control-localization-from-the-manifest)). Au lieu de cela, la localisation doit se produire lors de l’runtime à l’aide de blobs JSON distincts pour chaque paramètre régional. Nous vous suggérons d’utiliser une `switch` instruction qui teste [la propriété Office.context.displayLanguage](/javascript/api/office/office.context#office-office-context-displaylanguage-member). Voici un exemple.

```javascript
function GetContextualTabsJsonSupportedLocale () {
    var displayLanguage = Office.context.displayLanguage;

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

Ensuite, votre code appelle la fonction pour obtenir l’objet blob `requestCreateControls`local qui est transmis à , comme dans l’exemple suivant.

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>Meilleures pratiques pour les onglets contextuels personnalisés

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>Implémenter une autre expérience d’interface utilisateur lorsque les onglets contextuels personnalisés ne sont pas pris en charge

Certaines combinaisons de plateforme, Office application et de Office build ne sont pas prise en charge`requestCreateControls`. Votre add-in doit être conçu pour offrir une expérience de remplacement aux utilisateurs qui exécutent le module sur l’une de ces combinaisons. Les sections suivantes décrivent deux façons de fournir une expérience de retour.

#### <a name="use-noncontextual-tabs-or-controls"></a>Utiliser des onglets ou des contrôles nontexte

Il existe un élément manifeste, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), conçu pour créer une expérience de base dans un application qui implémente des onglets contextuels personnalisés lorsque le module est en cours d’exécution sur une application ou une plateforme qui ne prend pas en charge les onglets contextuels personnalisés.

La stratégie la plus simple pour utiliser cet élément consiste à définir un ou plusieurs onglets principaux personnalisés (c’est-à-dire, des onglets personnalisés *non* contextuels) dans le manifeste qui dupliquent les personnalisations du ruban des onglets contextuels personnalisés dans votre application. Toutefois, vous ajoutez `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` en tant que premier élément enfant des éléments [Group](../reference/manifest/group.md), [Control](../reference/manifest/control.md) et **Menu Item** dupliqués dans les onglets principaux personnalisés. L’effet de cette utilisation est le suivant :

- Si le add-in s’exécute sur une application et une plateforme qui prend en charge les onglets contextuels personnalisés, les groupes et contrôles principaux personnalisés n’apparaissent pas sur le ruban. Au lieu de cela, l’onglet contextuel personnalisé est créé lorsque le add-in appelle la `requestCreateControls` méthode.
- Si le add-in *s’exécute* `requestCreateControls`sur une application ou une plateforme qui ne prend pas en charge, les éléments apparaissent dans les onglets principaux personnalisés.

Voici un exemple. Notez que « MyButton » apparaît sur l’onglet principal personnalisé uniquement lorsque les onglets contextuels personnalisés ne sont pas pris en charge. Toutefois, le groupe parent et l’onglet principal personnalisé apparaissent, que les onglets contextuels personnalisés soient pris en charge ou non.

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

Pour plus d’exemples, [voir OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).

Lorsqu’un groupe parent ou un menu `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`est marqué avec, il n’est pas visible et tous ses marques enfants sont ignorés lorsque les onglets contextuels personnalisés ne sont pas pris en charge. Ainsi, peu importe si l’un de ces éléments enfants a l’élément **OverriddenByRibbonApi** ou sa valeur. En conséquence, si un élément de menu ou un contrôle doit être visible dans tous les contextes, non seulement il ne doit pas être marqué avec, `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`mais son *menu* ancêtre et son groupe ne doivent pas non plus être marqués de cette façon.

> [!IMPORTANT]
> Ne marquez pas *tous les* éléments enfants d’un groupe ou d’un menu avec `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`. Cela est inutile si l’élément parent est marqué pour des `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` raisons indiquées dans le paragraphe précédent. En outre, si vous laissez l’élément **OverriddenByRibbonApi** sur le parent ( `false`ou si vous le définissez sur ), le parent apparaît, que les onglets contextuels personnalisés soient pris en charge ou non, mais qu’ils soient vides lorsqu’ils sont pris en charge. Ainsi, si tous les éléments enfants ne doivent pas apparaître lorsque les onglets contextuels personnalisés sont pris en charge, marquez le parent avec `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>Utiliser des API qui montrent ou masquent un volet Des tâches dans des contextes spécifiés

En remplacement de **OverriddenByRibbonApi**, votre add-in peut définir un volet Des tâches avec des contrôles d’interface utilisateur qui dupliquent la fonctionnalité des contrôles dans un onglet contextuel personnalisé. Utilisez ensuite les [méthodes Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-showastaskpane-member(1)) et [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-hide-member(1)) pour afficher le volet Des tâches lorsque l’onglet contextuel aurait été affiché s’il était pris en charge. Pour plus d’informations sur l’utilisation de ces méthodes, voir Afficher ou masquer le volet Des tâches de [votre Office de travail](../develop/show-hide-add-in.md).

### <a name="handle-the-hostrestartneeded-error"></a>Gérer l’erreur HostRestartNeeded

Dans certains scénarios, Office ne peut pas mettre à jour le ruban et renvoie une erreur. Par exemple, si le complément est mis à niveau et que le complément mis à niveau dispose d'un autre groupe de commandes de complément personnalisé, l’application Office doit être fermée et ouverte de nouveau. La méthode `requestUpdate` renvoie l'erreur `HostRestartNeeded` jusqu'à ce que cela soit effectué. Votre code doit gérer cette erreur. Voici un exemple de comment. Dans ce cas, la méthode `reportError` affiche l’erreur à l’utilisateur.

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

- [Exemple de code : créer des onglets contextuels personnalisés sur le ruban](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-contextual-tabs)
- Community exemple d’onglets contextuels

> [!VIDEO https://www.youtube.com/embed/9tLfm4boQIo]