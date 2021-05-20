---
title: Créez des onglets contextuels personnalisés Office add-ins
description: Découvrez comment ajouter des onglets contextuels personnalisés à Office add-in.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: d03ac2c01c03353f3e2d1b54ba20616d7b42d93f
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555205"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a>Créez des onglets contextuels personnalisés Office add-ins

Un onglet contextuel est un contrôle d’onglet caché dans le ruban Office qui s’affiche dans la ligne d’onglet lorsqu’un événement spécifié se produit dans le document Office’affichage. Par exemple, **l’onglet Conception** de table qui apparaît sur Excel ruban lorsqu’une table est sélectionnée. Vous pouvez inclure des onglets contextuels personnalisés dans votre Office Add-in et spécifier quand ils sont visibles ou cachés, en créant des gestionnaires d’événements qui modifient la visibilité. (Toutefois, les onglets contextuels personnalisés ne répondent pas aux modifications de mise au point.)

> [!NOTE]
> Cet article suppose que vous connaissez la documentation décrite ci-après. Étudiez-la si vous n’avez pas récemment utilisé les commandes de complément (éléments de menu et boutons de ruban personnalisés).
>
> - [Concepts basiques pour les commandes de complément](add-in-commands.md)

> [!IMPORTANT]
> Les onglets contextuels personnalisés ne sont actuellement pris en charge Excel et uniquement sur ces plates-formes et builds :
>
> - Excel sur Windows (Microsoft 365 abonnement uniquement): Version 2102 (Build 13801.20294) ou plus tard.
> - Excel sur le web

> [!NOTE]
> Les onglets contextuels personnalisés ne fonctionnent que sur les plates-formes qui supporte les ensembles d’exigences suivants. Pour en savoir plus sur les ensembles d’exigences et la façon de travailler avec eux, [consultez spécifier Office applications et les exigences de l’API](../develop/specify-office-hosts-and-api-requirements.md).
>
> - [RibbonApi 1.2 RubanApi 1.2](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> Vous pouvez utiliser les vérifications de temps d’exécution dans votre code pour vérifier si l’hôte de l’utilisateur et la combinaison de plate-forme prend en charge ces ensembles d’exigences tels [que décrits dans spécifier les applications Office et les exigences de l’API](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code). (La technique de spécifier les ensembles d’exigences dans le manifeste, qui est également décrit dans cet article, ne fonctionne pas actuellement pour RibbonApi 1.2.) Alternativement, vous pouvez [implémenter une expérience d’interface utilisateur alternative lorsque les onglets contextuels personnalisés ne sont pas pris en charge](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

## <a name="behavior-of-custom-contextual-tabs"></a>Comportement des onglets contextuels personnalisés

L’expérience utilisateur des onglets contextuels personnalisés suit le modèle des onglets Office intégrés intégrés. Voici les principes de base pour les onglets contextuels personnalisés de placement :

- Lorsqu’un onglet contextuel personnalisé est visible, il apparaît à l’extrémité droite du ruban.
- Si un ou plusieurs onglets contextuels intégrés et un ou plusieurs onglets contextuels personnalisés provenant d’add-ins sont visibles en même temps, les onglets contextuels personnalisés sont toujours à droite de tous les onglets contextuels intégrés.
- Si votre module a plus d’un onglet contextuel et qu’il y a des contextes dans lesquels plus d’un est visible, ils apparaissent dans l’ordre dans lequel ils sont définis dans votre module. (La direction est dans la même direction que la langue Office; c’est-à-dire de gauche à droite dans les langues de gauche à droite, mais de droite à gauche dans les langues de droite à gauche.) Voir [Définir les groupes et les contrôles qui apparaissent sur l’onglet pour plus](#define-the-groups-and-controls-that-appear-on-the-tab) de détails sur la façon dont vous les définissez.
- Si plus d’un add-in a un onglet contextuel qui est visible dans un contexte spécifique, alors ils apparaissent dans l’ordre dans lequel les add-ins ont été lancés.
- Les *onglets* contextuels personnalisés, contrairement aux onglets de base personnalisés, ne sont pas ajoutés en permanence Office ruban de l’application. Ils ne sont présents que dans Office documents sur lesquels votre module est en cours d’exécution.

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>Étapes majeures pour inclure un onglet contextuel dans un module

Voici les principales étapes pour inclure un onglet contextuel personnalisé dans un module :

1. Configurez l’add-in pour utiliser un temps d’exécution partagé.
1. Définissez l’onglet et les groupes et contrôles qui y apparaissent.
1. Enregistrez l’onglet contextuel avec Office.
1. Spécifiez les circonstances dans lesquelles l’onglet sera visible.

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurez l’add-in pour utiliser un temps d’exécution partagé

L’ajout d’onglets contextuels personnalisés nécessite votre module d’utilisation pour utiliser le temps d’exécution partagé. Pour plus d’informations, [consultez Configurer un module d’accès pour utiliser un temps d’exécution partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>Définir les groupes et les contrôles qui apparaissent sur l’onglet

Contrairement aux onglets de base personnalisés, qui sont définis avec XML dans le manifeste, les onglets contextuels personnalisés sont définis à l’exécution avec un blob JSON. Votre code analyse le blob dans un objet JavaScript, puis passe l’objet [à la méthode Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) Les onglets contextuels personnalisés ne sont présents que dans les documents sur lesquels votre module est actuellement en cours d’exécution. Ceci est différent des onglets de base personnalisés qui sont ajoutés au ruban d’application Office lorsque l’add-in est installé et restent présents lorsqu’un autre document est ouvert. En outre, la `requestCreateControls` méthode ne peut être utilisée qu’une seule fois dans une session de votre module d’ajout. Si elle est appelée à nouveau, une erreur est lancée.

> [!NOTE]
> La structure des propriétés et des sous-propriétés du blob JSON (et des noms clés) est à peu près parallèle à la structure de [l’élément CustomTab](../reference/manifest/customtab.md) et de ses éléments descendants dans le manifeste XML.

Nous allons construire un exemple d’onglets contextuels JSON blob étape par étape. Le schéma complet de l’onglet contextuel JSON est [ àdynamic-ribbon.schema.jssur](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json). Si vous travaillez dans Visual Studio Code, vous pouvez utiliser ce fichier pour obtenir IntelliSense et valider votre JSON. Pour plus d’informations, [voir Édition JSON avec Visual Studio Code - Schémas et paramètres JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).


1. Commencez par créer une chaîne JSON avec deux propriétés de tableau nommées `actions` et `tabs` . Le `actions` tableau est une spécification de toutes les fonctions qui peuvent être exécutées par des contrôles sur l’onglet contextuel. Le `tabs` tableau définit un ou plusieurs onglets contextuels, *jusqu’à un maximum de 20*.

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. Cet exemple simple d’onglet contextuel n’aura qu’un seul bouton et, par conséquent, une seule action. Ajoutez ce qui suit en tant que seul membre du `actions` tableau. A propos de ce balisage, notez:

    - Les `id` propriétés et les propriétés sont `type` obligatoires.
    - La valeur de `type` peut être soit « ExecuteFunction » ou « ShowTaskpane ».
    - La `functionName` propriété n’est utilisée que lorsque la valeur `type` de est `ExecuteFunction` . C’est le nom d’une fonction définie dans le Fichier de fonction. Pour plus d’informations sur le Fichier de fonction, [consultez les concepts de base pour les commandes complémentaires](add-in-commands.md).
    - Dans une étape ultérieure, vous cartographiez cette action à un bouton sur l’onglet contextuel.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. Ajoutez ce qui suit en tant que seul membre du `tabs` tableau. A propos de ce balisage, notez:

    - La propriété `id` est requise. Utilisez un bref id descriptif unique parmi tous les onglets contextuels de votre module.
    - La propriété `label` est requise. Il s’agit d’une chaîne conviviale pour servir d’étiquette de l’onglet contextuel.
    - La propriété `groups` est requise. Il définit les groupes de contrôles qui apparaîtront sur l’onglet. Il doit avoir au moins un membre *et pas plus de 20*. (Il ya aussi des limites sur le nombre de contrôles que vous pouvez avoir sur un onglet contextuel personnalisé et qui limitera également le nombre de groupes que vous avez. Voir l’étape suivante pour plus d’informations.)

    > [!NOTE]
    > L’objet onglet peut également avoir une propriété `visible` optionnelle qui spécifie si l’onglet est visible immédiatement lorsque l’add-in démarre. Étant donné que les onglets contextuels sont normalement masqués jusqu’à ce qu’un événement utilisateur déclenche leur visibilité (comme l’utilisateur sélectionnant une entité d’un certain type dans le document), la propriété ne `visible` se présente `false` pas lorsqu’elle n’est pas présente. Dans une section ultérieure, nous montrons comment définir la propriété `true` en réponse à un événement.

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. Dans l’exemple continu simple, l’onglet contextuel n’a qu’un seul groupe. Ajoutez ce qui suit en tant que seul membre du `groups` tableau. A propos de ce balisage, notez:

    - Toutes les propriétés sont requises.
    - La `id` propriété doit être unique parmi tous les groupes de l’onglet. Utilisez une pièce d’identité brève et descriptive.
    - Il `label` s’agit d’une chaîne conviviale pour servir d’étiquette du groupe.
    - La `icon` valeur de la propriété est un tableau d’objets qui spécifient les icônes que le groupe aura sur le ruban en fonction de la taille du ruban et de la fenêtre d’application Office’application.
    - La `controls` valeur de la propriété est un éventail d’objets qui spécifient les boutons et les menus du groupe. Il doit y en avoir au moins un.

    > [!IMPORTANT]
    > *Le nombre total de contrôles sur l’ensemble de l’onglet ne peut pas être supérieur à 20.* Par exemple, vous pouvez avoir 3 groupes avec 6 contrôles chacun, et un quatrième groupe avec 2 contrôles, mais vous ne pouvez pas avoir 4 groupes avec 6 contrôles chacun.  

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

1. Chaque groupe doit avoir une icône d’au moins deux tailles, 32x32 px et 80x80 px. En option, vous pouvez également avoir des icônes de tailles 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, et 64x64 px. Office quelle icône utiliser en fonction de la taille du ruban et de la fenêtre d Office’application. Ajoutez les objets suivants au tableau d’icônes. (Si la taille de la fenêtre et du ruban est suffisamment grande pour qu’au moins une des *commandes* du groupe apparaisse, aucune icône de groupe n’apparaît. Par exemple, regardez le groupe **Styles sur** le ruban Word lorsque vous rétrécissez et élargissez la fenêtre Word.) A propos de ce balisage, notez:

    - Les deux propriétés sont requises.
    - `size`L’unité de mesure de propriété est pixels. Les icônes sont toujours carrées, de sorte que le nombre est à la fois la hauteur et la largeur.
    - La `sourceLocation` propriété spécifie l’URL complète de l’icône.

    > [!IMPORTANT]
    > Tout comme vous devez généralement modifier les URL dans le manifeste de l’add-in lorsque vous passez du développement à la production (comme changer le domaine de localhost à contoso.com), vous devez également modifier les URL dans vos onglets contextuels JSON.

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

1. Dans notre exemple continu simple, le groupe n’a qu’un seul bouton. Ajoutez l’objet suivant comme seul membre du `controls` tableau. A propos de ce balisage, notez:

    - Toutes les propriétés, sauf `enabled` , sont nécessaires.
    - `type` spécifie le type de contrôle. Les valeurs peuvent être « Bouton », « Menu » ou « MobileButton ».
    - `id` peut être jusqu’à 125 caractères. 
    - `actionId` doit être l’ID d’une action définie dans le `actions` tableau. (Voir l’étape 1 de cette section.)
    - `label` est une chaîne conviviale pour servir d’étiquette du bouton.
    - `superTip` représente une forme riche de pointe d’outil. Les propriétés `title` et les propriétés sont `description` requises.
    - `icon` spécifie les icônes pour le bouton. Les remarques précédentes sur l’icône du groupe s’appliquent ici aussi.
    - `enabled` (facultatif) précise si le bouton est activé lorsque l’onglet contextuel apparaît démarre. La valeur par défaut si elle n’est pas présente est `true` . 

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>Enregistrez l’onglet contextuel Office avec requestCreateControls

L’onglet contextuel est enregistré Office en [appelant la méthode Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) Cela se fait généralement soit dans la fonction qui est assignée `Office.initialize` à ou avec la `Office.onReady` méthode. Pour en savoir plus sur ces méthodes et parasinant l’add-in, [voir Initialiser Office add-in](../develop/initialize-add-in.md). Vous pouvez toutefois appeler la méthode à tout moment après l’initialisation.

> [!IMPORTANT]
> La `requestCreateControls` méthode ne peut être appelée qu’une seule fois dans une session donnée d’un add-in. Une erreur est lancée si elle est appelée à nouveau.

Voici un exemple. Notez que la chaîne JSON doit être convertie en objet JavaScript avec la méthode `JSON.parse` avant qu’elle puisse être transmise à une fonction JavaScript.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>Spécifiez les contextes lorsque l’onglet sera visible avec requestUpdate

En règle générale, un onglet contextuel personnalisé doit apparaître lorsqu’un événement initié par l’utilisateur modifie le contexte d’ajout. Considérez un scénario dans lequel l’onglet doit être visible lorsque, et seulement quand, un graphique (sur la feuille de travail par défaut d’un Excel de travail) est activé.

Commencez par affecter des gestionnaires. Cela est généralement fait dans la `Office.onReady` méthode comme dans l’exemple suivant qui attribue les gestionnaires (créés dans une étape ultérieure) aux événements et aux événements de tous les `onActivated` graphiques de la feuille de `onDeactivated` travail.

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

Ensuite, définissez les gestionnaires. Ce qui suit est un exemple simple d’un `showDataTab` , mais voir Manipulation de [l’erreur HostRestartNeeded plus](#handle-the-hostrestartneeded-error) tard dans cet article pour une version plus robuste de la fonction. Tenez compte du code suivant :

- Office effectue un contrôle lorsqu’il met à jour l’état du ruban. La [méthode Office.ribbon.requestUpdate fait](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) la queue pour une demande de mise à jour. La méthode résoudra `Promise` l’objet dès qu’il aura mis la demande en file d’attente, et non lorsque le ruban sera mis à jour.
- Le paramètre de `requestUpdate` la méthode est un objet [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) qui (1) spécifie l’onglet par son ID exactement *comme spécifié dans le JSON* et (2) spécifie la visibilité de l’onglet.
- Si vous avez plus d’un onglet contextuel personnalisé qui devrait être visible dans le même contexte, il vous suffit d’ajouter des objets d’onglet supplémentaires au `tabs` tableau.

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

Le gestionnaire pour cacher l’onglet est presque identique, sauf qu’il définit la `visible` propriété de nouveau à `false` .

La Office javascript fournit également plusieurs interfaces (types) pour faciliter la construction de `RibbonUpdateData` l’objet. Ce qui suit est `showDataTab` la fonction dans TypeScript et il fait usage de ces types.

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Visibilité de l’onglet Basculement et état activé d’un bouton en même temps

La `requestUpdate` méthode est également utilisée pour basculer l’état activé ou désactivé d’un bouton personnalisé sur un onglet contextuel personnalisé ou un onglet de base personnalisé. Pour plus de détails à ce [sujet, consultez activer et désactiver les commandes add-in](disable-add-in-commands.md). Il peut y avoir des scénarios dans lesquels vous souhaitez modifier à la fois la visibilité d’un onglet et l’état activé d’un bouton en même temps. Vous pouvez le faire avec un seul appel de `requestUpdate` . Ce qui suit est un exemple dans lequel un bouton sur un onglet de base est activé en même temps qu’un onglet contextuel est rendu visible.

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

Dans l’exemple suivant, le bouton activé est sur le même onglet contextuel qui est rendu visible.

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

## <a name="localizing-the-json-blob"></a>Localisation du blob JSON

Le blob JSON qui est transmis n’est `requestCreateControls` pas localisé de la même manière que le balisage manifeste pour les onglets de base personnalisés est localisé (qui est décrit à la localisation de contrôle à partir du [manifeste).](../develop/localization.md#control-localization-from-the-manifest) Au lieu de cela, la localisation doit se produire à l’heure d’exécution en utilisant des blobs JSON distincts pour chaque lieu. Nous vous suggérons `switch` d’utiliser une instruction qui [teste Office.context.display Propriété](/javascript/api/office/office.context#displayLanguage) de la langue. Voici un exemple :

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

Ensuite, votre code appelle la fonction pour obtenir le blob localisé qui est transmis `requestCreateControls` à , comme dans l’exemple suivant:

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>Meilleures pratiques pour les onglets contextuels personnalisés

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>Implémentez une expérience d’interface utilisateur alternative lorsque les onglets contextuels personnalisés ne sont pas pris en charge

Certaines combinaisons de plate-forme, Office’application et Office construire ne supportent pas `requestCreateControls` . Votre module d’ajout doit être conçu pour offrir une expérience alternative aux utilisateurs qui font fonctionner l’add-in sur l’une de ces combinaisons. Les sections suivantes décrivent deux façons d’offrir une expérience de repli.

#### <a name="use-noncontextual-tabs-or-controls"></a>Utilisez des onglets ou des commandes non textuels

Il ya un élément manifeste, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), qui est conçu pour créer une expérience de secours dans un add-in qui implémente des onglets contextuels personnalisés lorsque l’add-in est en cours d’exécution sur une application ou une plate-forme qui ne prend pas en charge les onglets contextuels personnalisés. 

La stratégie la plus simple pour l’utilisation de cet élément est que vous définissez dans le manifeste un ou plusieurs onglets de base personnalisés *(c’est-à-dire* des onglets personnalisés non textuels) qui dupliquent les personnalisations de ruban des onglets contextuels personnalisés de votre module. Mais vous ajoutez `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` comme premier élément enfant du [CustomTab](../reference/manifest/customtab.md). L’effet de cette chose est le suivant :

- Si l’add-in s’exécute sur une application et une plate-forme qui supporte les onglets contextuels personnalisés, alors l’onglet de base personnalisé n’apparaîtra pas sur le ruban. Au lieu de cela, l’onglet contextuel personnalisé sera créé lorsque l’add-in appelle la `requestCreateControls` méthode.
- Si *l’add-in s’exécute* sur une application ou une plate-forme qui ne prend pas en `requestCreateControls` charge, alors l’onglet de base personnalisé apparaît sur le ruban.

Ce qui suit est un exemple de cette stratégie simple.

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

Cette stratégie simple utilise un onglet de base personnalisé qui reflète un onglet contextuel personnalisé avec ses groupes d’enfants et ses contrôles, mais vous pouvez utiliser une stratégie plus complexe. `<OverriddenByRibbonApi>`L’élément peut également être ajouté en tant que (premier) élément enfant aux éléments [groupe](../reference/manifest/group.md) [et contrôle](../reference/manifest/control.md) [(type de bouton et](../reference/manifest/control.md#button-control) type de [menu)](../reference/manifest/control.md#menu-dropdown-button-controls)et éléments de `<Item>` menu. Ce fait vous permet de distribuer les groupes et les contrôles qui apparaîtraient autrement sur l’onglet contextuel entre différents groupes, boutons et menus dans divers onglets de base personnalisés. Voici un exemple. Notez que « MyButton » n’apparaîtra sur l’onglet de base personnalisé que lorsque les onglets contextuels personnalisés ne sont pas pris en charge. Mais le groupe parent et l’onglet de base personnalisé s’afficheront indépendamment du fait que les onglets contextuels personnalisés soient pris en charge.

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

Lorsqu’un onglet parent, un groupe ou un menu est marqué `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` avec, alors il n’est pas visible, et tout son balisage enfant est ignoré, lorsque les onglets contextuels personnalisés ne sont pas pris en charge. Donc, il n’a pas d’importance si l’un de ces éléments enfant ont `<OverriddenByRibbonApi>` l’élément ou ce que sa valeur est. L’implication de ceci est que si un élément de menu, le contrôle, ou le groupe doit être visible dans tous les contextes, alors non seulement devrait-il pas être marqué `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` avec, mais *son menu d’ancêtre, groupe, et onglet ne doit pas non plus être marqué de cette façon.*

> [!IMPORTANT]
> Ne marquez pas *tous les* éléments enfant d’un onglet, d’un groupe ou d’un menu avec `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` . Cela ne va pas si l’élément parent est marqué pour `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` des raisons données dans le paragraphe précédent. En outre, si vous laissez de côté `<OverriddenByRibbonApi>` le sur le parent (ou le définir à ), alors le parent apparaîtra indépendamment du fait que les `false` onglets contextuels personnalisés sont pris en charge, mais il sera vide quand ils sont pris en charge. Ainsi, si tous les éléments enfant ne devraient pas apparaître lorsque les onglets contextuels personnalisés sont pris en charge, marquez le parent, et seulement le parent, avec `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>Utilisez des API qui affichent ou cachent un volet de tâche dans des contextes spécifiques

Comme alternative à , votre add-in peut définir un volet de tâche avec des contrôles `<OverriddenByRibbonApi>` d’interface utilisateur qui dupliquent la fonctionnalité des contrôles sur un onglet contextuel personnalisé. Ensuite, utilisez [les méthodes Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) [et Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) pour afficher le volet de tâche quand, et seulement quand, l’onglet contextuel aurait été affiché s’il avait été pris en charge. Pour plus de détails sur la façon d’utiliser ces [méthodes, voir Afficher ou masquer le volet de tâche de votre Office Add-in](../develop/show-hide-add-in.md).

### <a name="handle-the-hostrestartneeded-error"></a>Gérer l’erreur HostRestartNeeded

Dans certains scénarios, Office ne peut pas mettre à jour le ruban et renvoie une erreur. Par exemple, si le complément est mis à niveau et que le complément mis à niveau dispose d'un autre groupe de commandes de complément personnalisé, l’application Office doit être fermée et ouverte de nouveau. La méthode `requestUpdate` renvoie l'erreur `HostRestartNeeded` jusqu'à ce que cela soit effectué. Votre code doit gérer cette erreur. Ce qui suit est un exemple de la façon dont. Dans ce cas, la méthode `reportError` affiche l’erreur à l’utilisateur.

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
