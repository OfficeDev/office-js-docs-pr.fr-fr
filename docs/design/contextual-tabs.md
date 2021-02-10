---
title: Créer des onglets contextuels personnalisés dans les add-ins Office
description: Découvrez comment ajouter des onglets contextuels personnalisés à votre add-in Office.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 67588e04d6ea95bc581c51e274c8135cfa5afd50
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173919"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a>Créer des onglets contextuels personnalisés dans des compléments Office (préversion)

Un onglet contextuel est un contrôle onglet masqué dans le ruban Office qui s’affiche dans la ligne d’onglet lorsqu’un événement spécifié se produit dans le document Office. Par exemple, **l’onglet Création de** tableau qui apparaît sur le ruban Excel lorsqu’un tableau est sélectionné. Vous pouvez inclure des onglets contextuels personnalisés dans votre add-in Office et spécifier quand ils sont visibles ou masqués en créant des handlers d’événements qui modifient la visibilité. (Toutefois, les onglets contextuels personnalisés ne répondent pas aux changements de focus.)

> [!NOTE]
> Cet article suppose que vous connaissez la documentation décrite ci-après. Étudiez-la si vous n’avez pas récemment utilisé les commandes de complément (éléments de menu et boutons de ruban personnalisés).
>
> - [Concepts basiques pour les commandes de complément](add-in-commands.md)

> [!IMPORTANT]
> Les onglets contextuels personnalisés sont en prévisualisation. Testez-les dans un environnement de développement ou de test, mais ne les ajoutez pas à un module de production.
>
> Les onglets contextuels personnalisés sont actuellement uniquement pris en charge sur Excel et uniquement sur ces plateformes et builds :
>
> - Excel sur Windows (Microsoft 365 uniquement, et non la licence perpétuelle) : version 2011 (build 13426.20274). Votre abonnement Microsoft 365 doit peut-être être sur le canal actuel [(prévisualisation)](https://insider.office.com/join/windows) anciennement appelé « Canal mensuel (ciblé) » ou « Insider Slow ».

> [!NOTE]
> Les onglets contextuels personnalisés fonctionnent uniquement sur les plateformes qui supportent les ensembles de conditions requises suivants. Pour plus d’informations sur les ensembles de conditions requises et sur leur utilisation, voir Spécifier les [applications Office et les conditions requises des API.](../develop/specify-office-hosts-and-api-requirements.md)
>
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a>Comportement des onglets contextuels personnalisés

L’expérience utilisateur pour les onglets contextuels personnalisés suit le modèle des onglets contextuels Office intégrés. Les principes de base des onglets contextuels personnalisés de placement sont les suivants :

- Lorsqu’un onglet contextuel personnalisé est visible, il apparaît à l’extrémité droite du ruban.
- Si un ou plusieurs onglets contextuels intégrés et un ou plusieurs onglets contextuels personnalisés des modules sont visibles en même temps, les onglets contextuels personnalisés sont toujours à droite de tous les onglets contextuels intégrés.
- Si votre add-in possède plusieurs onglets contextuels et qu’il existe des contextes dans lesquels plusieurs onglets sont visibles, ils apparaissent dans l’ordre dans lequel ils sont définis dans votre module. (Le sens est identique à celui de la langue d’Office ; c’est-à-dire, de gauche à droite dans les langues de gauche à droite, mais de droite à gauche dans les langues de droite à gauche.) Pour [plus d’informations sur](#define-the-groups-and-controls-that-appear-on-the-tab) leur définition, voir Définir les groupes et les contrôles qui apparaissent sous l’onglet.
- Si plusieurs d’entre eux ont un onglet contextuel visible dans un contexte spécifique, ils apparaissent dans l’ordre dans lequel les modules ont été lancés.
- Contrairement *aux* onglets principaux personnalisés, les onglets contextuels personnalisés ne sont pas ajoutés définitivement au ruban de l’application Office. Elles sont présentes uniquement dans les documents Office sur lesquels votre module est en cours d’exécution.

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>Étapes principales pour l’ajout d’un onglet contextuel dans un add-in

Les étapes principales d’ajout d’un onglet contextuel personnalisé dans un add-in sont les suivantes :

1. Configurez le add-in pour utiliser un runtime partagé.
1. Définissez l’onglet, ainsi que les groupes et les contrôles qui y apparaissent.
1. Inscrivez l’onglet contextuel auprès d’Office.
1. Spécifiez les circonstances dans le cas où l’onglet sera visible.

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurer le add-in pour utiliser un runtime partagé

L’ajout d’onglets contextuels personnalisés nécessite que votre add-in utilise le runtime partagé. Pour plus d’informations, [voir Configurer un module complémentaire pour utiliser un runtime partagé.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>Définir les groupes et les contrôles qui apparaissent sous l’onglet

Contrairement aux onglets principaux personnalisés, qui sont définis avec du XML dans le manifeste, les onglets contextuels personnalisés sont définis lors de l’runtime avec un blob JSON. Votre code parse le blob dans un objet JavaScript, puis passe l’objet à la méthode [Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) Les onglets contextuels personnalisés sont uniquement présents dans les documents sur lesquels votre module est en cours d’exécution. Cela est différent des onglets principaux personnalisés qui sont ajoutés au ruban de l’application Office lorsque le module est installé et restent présents à l’ouverture d’un autre document. En outre, `requestCreateControls` la méthode ne peut être exécuté qu’une seule fois dans une session de votre add-in. Si elle est appelée à nouveau, une erreur est lancée.

> [!NOTE]
> La structure des propriétés et sous-propriétés de l’objet blob JSON (et les noms clés) est à peu près parallèle à la structure de l’élément [CustomTab](../reference/manifest/customtab.md) et de ses éléments descendants dans le manifeste XML.

Nous allons créer un exemple d’objet blob JSON onglets contextuel pas à pas. (Le schéma complet de l’onglet contextuel JSON est [dynamic-ribbon.schema.jssur](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json). Il se peut que ce lien ne fonctionne pas pendant la période d’aperçu pour les onglets contextuels. Si le lien ne fonctionne pas, vous pouvez trouver le dernier brouillon du schéma à [l'dynamic-ribbon.schema.jssur](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon/1.0/dynamic-ribbon.schema.json).) Si vous travaillez dans Visual Studio Code, vous pouvez utiliser ce fichier pour obtenir IntelliSense et valider votre JSON. Pour plus d’informations, voir [Modification de JSON avec Visual Studio Code - Schémas et paramètres JSON.](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)


1. Commencez par créer une chaîne JSON avec deux propriétés de tableau `actions` nommées et `tabs` . Le tableau est une spécification de toutes les fonctions qui peuvent être exécutées par des `actions` contrôles sous l’onglet contextuel. Le `tabs` tableau définit un ou plusieurs onglets contextuels, *jusqu’à un maximum de 20*.

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. Cet exemple simple d’onglet contextuel n’aura qu’un seul bouton et, par conséquent, une seule action. Ajoutez ce qui suit en tant que seul membre du `actions` tableau. À propos de ce markup, notez :

    - Les `id` `type` propriétés et les propriétés sont obligatoires.
    - La valeur `type` peut être « ExecuteFunction » ou « ShowTaskpane ».
    - La `functionName` propriété est utilisée uniquement lorsque la valeur est `type` `ExecuteFunction` . Il s’agit du nom d’une fonction définie dans functionFile. Pour plus d’informations sur FunctionFile, voir [Concepts de base pour les commandes de module complémentaire.](add-in-commands.md)
    - Dans une étape ultérieure, vous allez ma cartographier cette action sur un bouton de l’onglet contextuel.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. Ajoutez ce qui suit en tant que seul membre du `tabs` tableau. À propos de ce markup, notez :

    - La propriété `id` est requise. Utilisez un bref ID descriptif unique parmi tous les onglets contextuels de votre add-in.
    - La propriété `label` est requise. Il s’agit d’une chaîne conviviale qui sert d’étiquette à l’onglet contextuel.
    - La propriété `groups` est requise. Il définit les groupes de contrôles qui apparaîtront sous l’onglet. Elle doit avoir au moins un *membre et pas plus de 20*. (Il existe également des limites au nombre de contrôles que vous pouvez avoir sur un onglet contextuel personnalisé et qui limitent également le nombre de groupes que vous avez. Pour plus d’informations, voir l’étape suivante.)

    > [!NOTE]
    > L’objet tabulation peut également avoir une propriété facultative qui spécifie si l’onglet est visible immédiatement au démarrage `visible` du module. Dans la mesure où les onglets contextuels sont normalement masqués jusqu’à ce qu’un événement utilisateur déclenche leur visibilité (par exemple, l’utilisateur sélectionnant une entité d’un type dans le document), la propriété se présente par défaut lorsqu’elle n’est pas `visible` `false` présente. Dans une section ultérieure, nous montrons comment définir la propriété en réponse `true` à un événement.

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
    - La `id` propriété doit être unique parmi tous les groupes de l’onglet. Utilisez un ID bref et descriptif.
    - Il `label` s’agit d’une chaîne conviviale qui sert d’étiquette au groupe.
    - La valeur de la propriété est un tableau d’objets qui spécifient les icônes que le groupe aura sur le ruban en fonction de la taille du ruban et de la fenêtre de `icon` l’application Office.
    - La valeur de la propriété est un tableau d’objets qui spécifient les boutons et `controls` les menus du groupe. Il doit y en avoir au moins un.

    > [!IMPORTANT]
    > *Le nombre total de contrôles sur l’onglet entier ne peut pas être supérieur à 20.* Par exemple, vous pouvez avoir 3 groupes avec 6 contrôles chacun et un quatrième groupe avec 2 contrôles, mais vous ne pouvez pas avoir 4 groupes avec 6 contrôles chacun.  

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

1. Chaque groupe doit avoir une icône d’au moins deux tailles, 32 x 32 px et 80 x 80 px. Si vous le souhaitez, vous pouvez également avoir des icônes de tailles 16 x 16 px, 20 x 20 px, 24 x 24 px, 40 x 40 px, 48 x 48 px et 64 x 64 px. Office décide de l’icône à utiliser en fonction de la taille du ruban et de la fenêtre de l’application Office. Ajoutez les objets suivants au tableau d’icônes. (Si les tailles de la fenêtre et  du ruban sont suffisamment grandes pour qu’au moins l’un des contrôles du groupe apparaisse, aucune icône de groupe ne s’affiche. Pour obtenir un exemple, regardez le groupe **Styles** sur le ruban Word lorsque vous réduirez et développez la fenêtre Word.) À propos de ce markup, notez :

    - Les deux propriétés sont obligatoires.
    - `size`L’unité de mesure de propriété est pixels. Les icônes sont toujours carrées, de sorte que le nombre est à la fois la hauteur et la largeur.
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

    - Toutes les propriétés, à l’exception `enabled` de , sont obligatoires.
    - `type` spécifie le type de contrôle. Les valeurs peuvent être « Button », « Menu » ou « MobileButton ».
    - `id` peut prendre jusqu’à 125 caractères. 
    - `actionId` doit être l’ID d’une action définie dans le `actions` tableau. (Voir l’étape 1 de cette section.)
    - `label` est une chaîne conviviale qui sert d’étiquette au bouton.
    - `superTip` représente une forme enrichie d’info-conseil. Les `title` propriétés et les `description` propriétés sont requises.
    - `icon` spécifie les icônes du bouton. Les remarques précédentes sur l’icône de groupe s’appliquent également ici.
    - `enabled` (facultatif) indique si le bouton est activé au démarrage de l’onglet contextuel. La valeur par défaut, si elle n’est pas présente, est `true` . 

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
 
Voici l’exemple complet du blob JSON :

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

L’onglet contextuel est inscrit auprès d’Office en appelant [la méthode Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) Cette tâche est généralement effectuée dans la fonction affectée à la méthode ou `Office.initialize` avec `Office.onReady` celle-ci. Pour plus d’informations sur ces méthodes et l’initialisation du add-in, voir [Initialize your Office Add-in](../develop/initialize-add-in.md). Vous pouvez toutefois appeler la méthode à tout moment après l’initialisation.

> [!IMPORTANT]
> La `requestCreateControls` méthode ne peut être appelée qu’une seule fois dans une session donnée d’un add-in. Une erreur est lancée si elle est appelée à nouveau.

Voici un exemple. Notez que la chaîne JSON doit être convertie en objet JavaScript avec la méthode pour pouvoir être transmise `JSON.parse` à une fonction JavaScript.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>Spécifier les contextes où l’onglet sera visible avec requestUpdate

En règle générale, un onglet contextuel personnalisé doit apparaître lorsqu’un événement initié par l’utilisateur modifie le contexte du add-in. Envisagez un scénario dans lequel l’onglet doit être visible lorsque, et uniquement quand, un graphique (dans la feuille de calcul par défaut d’un workbook Excel) est activé.

Commencez par affecter des handlers. Cela est généralement effectué dans la méthode comme dans l’exemple suivant qui affecte des handlers (créés à une étape ultérieure) aux événements et aux graphiques de la feuille `Office.onReady` `onActivated` de `onDeactivated` calcul.

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

Ensuite, définissez les handlers. Voici un exemple simple d’une erreur, mais voir Gestion de l’erreur `showDataTab` [HostRestartNeeded](#handle-the-hostrestartneeded-error) plus loin dans cet article pour obtenir une version plus robuste de la fonction. Tenez compte du code suivant :

- Office effectue un contrôle lorsqu’il met à jour l’état du ruban. La  [méthode Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) met en file d’attente une demande de mise à jour. La méthode résout l’objet dès qu’il a mis la demande en file d’attente, et non lorsque `Promise` le ruban est réellement mis à jour.
- Le paramètre de la méthode est un objet `requestUpdate` [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) qui (1) spécifie l’onglet par son ID exactement comme spécifié dans le *JSON* et (2) spécifie la visibilité de l’onglet.
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

Le handler pour masquer l’onglet est presque identique, sauf qu’il définit à `visible` nouveau la propriété sur `false` .

La bibliothèque JavaScript Office fournit également plusieurs interfaces (types) pour faciliter la construction de `RibbonUpdateData` l’objet. Voici la fonction dans TypeScript qui utilise `showDataTab` ces types.

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Activer la visibilité de l’onglet et l’état activé d’un bouton en même temps

La méthode est également utilisée pour activer ou désactiver l’état d’un bouton personnalisé sur un onglet contextuel personnalisé ou un `requestUpdate` onglet principal personnalisé. Pour plus d’informations à ce sujet, voir [Enable and Disable Add-in Commands](disable-add-in-commands.md). Il peut y avoir des scénarios dans lesquels vous souhaitez modifier la visibilité d’un onglet et l’état activé d’un bouton en même temps. Vous pouvez le faire avec un seul appel de `requestUpdate` . Voici un exemple dans lequel un bouton d’un onglet principal est activé en même temps qu’un onglet contextuel est rendu visible.

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

## <a name="localizing-the-json-blob"></a>Localisation de l’objet blob JSON

Le blob JSON passé à n’est pas localisée de la même façon que le marques de manifeste pour les onglets principaux personnalisés est localisée (ce qui est décrit lors de la localisation du contrôle à partir du `requestCreateControls` [manifeste).](../develop/localization.md#control-localization-from-the-manifest) Au lieu de cela, la localisation doit se produire lors de l’runtime à l’aide de blobs JSON distincts pour chaque paramètre régional. Nous vous suggérons d’utiliser `switch` une instruction qui teste la propriété [Office.context.displayLanguage.](/javascript/api/office/office.context#displayLanguage) Voici un exemple :

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

Ensuite, votre code appelle la fonction pour obtenir l’objet blob local qui est transmis à , comme `requestCreateControls` dans l’exemple suivant :

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>Meilleures pratiques pour les onglets contextuels personnalisés

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>Implémenter une autre expérience d’interface utilisateur lorsque les onglets contextuels personnalisés ne sont pas pris en charge

Certaines combinaisons de plateforme, d’application Office et de build Office ne sont pas prise en `requestCreateControls` charge. Votre add-in doit être conçu pour offrir une expérience de remplacement aux utilisateurs qui exécutent le module sur l’une de ces combinaisons. Les sections suivantes décrivent deux façons de fournir une expérience de retour.

#### <a name="use-noncontextual-tabs-or-controls"></a>Utiliser des onglets ou des contrôles nontexte

Il existe un élément manifeste, [OverriddenByRibbonApi,](../reference/manifest/overriddenbyribbonapi.md)conçu pour créer une expérience de base dans un application qui implémente des onglets contextuels personnalisés lorsque le module est en cours d’exécution sur une application ou une plateforme qui ne prend pas en charge les onglets contextuels personnalisés. 

La stratégie la plus simple pour l’utilisation de cet élément est que vous définissez dans le manifeste un ou plusieurs onglets principaux personnalisés (c’est-à-dire, des onglets personnalisés *non* contextuels) qui dupliquent les personnalisations du ruban des onglets contextuels personnalisés dans votre application. Mais vous `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` ajoutez en tant que premier élément enfant de [CustomTab](../reference/manifest/customtab.md). L’effet de cette utilisation est le suivant :

- Si le add-in s’exécute sur une application et une plateforme qui prend en charge les onglets contextuels personnalisés, l’onglet principal personnalisé n’apparaît pas sur le ruban. Au lieu de cela, l’onglet contextuel personnalisé est créé lorsque le add-in appelle la `requestCreateControls` méthode.
- Si le add-in *s’exécute* sur une application ou une plateforme qui ne prend pas en charge, l’onglet principal personnalisé `requestCreateControls` apparaît sur le ruban.

Voici un exemple de cette stratégie simple.

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
              <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  ...
                  <Action ...>
...
</OfficeApp>
```

Cette stratégie simple utilise un onglet principal personnalisé qui met en miroir un onglet contextuel personnalisé avec ses groupes et ses contrôles enfants, mais vous pouvez utiliser une stratégie plus complexe. L’élément peut également être ajouté en tant que (premier) élément enfant aux éléments Group et Control (type de bouton et type de `<OverriddenByRibbonApi>` [menu)](../reference/manifest/control.md#menu-dropdown-button-controls)et éléments de [](../reference/manifest/group.md) [](../reference/manifest/control.md) [](../reference/manifest/control.md#button-control) `<Item>` menu. Cela vous permet de distribuer les groupes et les contrôles qui apparaîtraient dans l’onglet contextuel entre différents groupes, boutons et menus dans différents onglets principaux personnalisés. Voici un exemple. Notez que « MyButton » apparaît sur l’onglet principal personnalisé uniquement lorsque les onglets contextuels personnalisés ne sont pas pris en charge. Toutefois, le groupe parent et l’onglet principal personnalisé apparaissent, que les onglets contextuels personnalisés soient pris en charge ou non.

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
                <Control ... id="MyButton">
                  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
                  ...
                  <Action ...>
...
</OfficeApp>
```

Pour plus d’exemples, [voir OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).

Lorsqu’un onglet, un groupe ou un menu parent est marqué avec, il n’est pas visible et tout son markup enfant est ignoré, lorsque les onglets contextuels personnalisés ne sont pas pris en `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` charge. Ainsi, peu importe si l’un de ces éléments enfants a l’élément ou `<OverriddenByRibbonApi>` sa valeur. En conséquence, si un élément de menu, un contrôle ou un groupe doit être visible dans tous les contextes, non seulement il ne doit pas être marqué avec, mais son `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` *ancêtre menu,* groupe et onglet ne doit pas non plus être marqué de cette façon.

> [!IMPORTANT]
> Ne marquez pas *tous les* éléments enfants d’un onglet, d’un groupe ou d’un menu avec `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` . Cela est inutile si l’élément parent est marqué pour des `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` raisons indiquées dans le paragraphe précédent. En outre, si vous ne le faites pas sur le parent (ou si vous le définissez sur ), le parent apparaît, que les onglets contextuels personnalisés soient pris en charge ou non, mais qu’ils soient vides lorsqu’ils sont pris en `<OverriddenByRibbonApi>` `false` charge. Ainsi, si tous les éléments enfants ne doivent pas apparaître lorsque les onglets contextuels personnalisés sont pris en charge, marquez le parent et uniquement le parent, avec `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>Utiliser des API qui montrent ou masquent un volet Des tâches dans des contextes spécifiés

En remplacement, votre add-in peut définir un volet Des tâches avec des contrôles d’interface utilisateur qui dupliquent les fonctionnalités des contrôles sur un `<OverriddenByRibbonApi>` onglet contextuel personnalisé. Utilisez ensuite les méthodes [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) et [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) pour afficher le volet Office quand, et uniquement quand, l’onglet contextuel aurait été affiché s’il était pris en charge. Pour plus d’informations sur l’utilisation de ces méthodes, voir Afficher ou masquer le volet Office de [votre add-in Office.](../develop/show-hide-add-in.md)

### <a name="handle-the-hostrestartneeded-error"></a>Gérer l’erreur HostRestartNeeded

Dans certains scénarios, Office ne peut pas mettre à jour le ruban et renvoie une erreur. Par exemple, si le complément est mis à niveau et que le complément mis à niveau dispose d'un autre groupe de commandes de complément personnalisé, l’application Office doit être fermée et ouverte de nouveau. La méthode `requestUpdate` renvoie l'erreur `HostRestartNeeded` jusqu'à ce que cela soit effectué. Votre code doit gérer cette erreur. Voici un exemple de comment. Dans ce cas, la méthode `reportError` affiche l’erreur à l’utilisateur.

```javascript
function showDataTab() {
    try {
        await Office.ribbon.requestUpdate({
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
