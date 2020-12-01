---
title: Créer des onglets contextuels personnalisés dans les compléments Office
description: Découvrez comment ajouter des onglets contextuels personnalisés à votre complément Office.
ms.date: 11/20/2020
localization_priority: Normal
ms.openlocfilehash: 49a773aca0651b88c972c24a4cde0aa1e300d5e7
ms.sourcegitcommit: 6619e07cdfa68f9fa985febd5f03caf7aee57d5e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/30/2020
ms.locfileid: "49505553"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a>Créer des onglets contextuels personnalisés dans les compléments Office (aperçu)

Un onglet contextuel est un contrôle d’onglet masqué dans le ruban Office qui s’affiche dans la ligne d’onglet lorsqu’un événement spécifié se produit dans le document Office. Par exemple, l’onglet **création de table** qui s’affiche dans le ruban Excel lorsqu’un tableau est sélectionné. Vous pouvez inclure des onglets contextuels personnalisés dans votre complément Office et spécifier lorsqu’ils sont visibles ou masqués, en créant des gestionnaires d’événements qui modifient la visibilité. (Toutefois, les onglets contextuels personnalisés ne répondent pas aux changements de focus.)

> [!NOTE]
> Cet article suppose que vous connaissez la documentation décrite ci-après. Étudiez-la si vous n’avez pas récemment utilisé les commandes de complément (éléments de menu et boutons de ruban personnalisés).
>
> - [Concepts basiques pour les commandes de complément](add-in-commands.md)

> [!IMPORTANT]
> Les onglets contextuels personnalisés sont en aperçu. Essayez de les tester dans un environnement de développement ou de test, mais ne les ajoutez pas à un complément de production.
>
> Les onglets contextuels personnalisés sont actuellement uniquement pris en charge sur Excel et uniquement sur ces plateformes et génèrent les éléments suivants :
>
> - Excel sur Windows (Microsoft 365 uniquement, pas une licence perpétuelle) : version 2011 (Build 13426,20274). Votre abonnement Microsoft 365 doit peut-être être sur le [canal actuel (](https://insider.office.com/join/windows) préversion) précédemment appelé « canal mensuel (ciblé) » ou « Insider Slower ».

> [!NOTE]
> Les onglets contextuels personnalisés fonctionnent uniquement sur les plateformes qui prennent en charge les ensembles de conditions requises suivants. Pour plus d’informations sur les ensembles de conditions requises et la façon de les utiliser, voir [spécifier les applications Office et les conditions requises](../develop/specify-office-hosts-and-api-requirements.md)pour les API.
>
> - [SharedRuntime 1,1](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a>Comportement des onglets contextuels personnalisés

L’expérience utilisateur pour les onglets contextuels personnalisés suit le modèle des onglets contextuels Office intégrés. Les éléments suivants sont les principes de base pour les onglets contextuels personnalisés de placement :

- Lorsqu’un onglet contextuel personnalisé est visible, il s’affiche à l’extrémité droite du ruban.
- Si un ou plusieurs onglets contextuels prédéfinis et un ou plusieurs onglets contextuels personnalisés provenant de compléments sont visibles en même temps, les onglets contextuels personnalisés sont toujours à droite de tous les onglets contextuels intégrés.
- Si votre complément comporte plusieurs onglets contextuels et des contextes dans lesquels plusieurs sont visibles, ils apparaissent dans l’ordre dans lequel ils sont définis dans votre complément. (Le sens est le même que celui de la langue d’Office ; autrement dit, de gauche à droite dans les langues se lisant de gauche à droite, mais de droite à gauche dans les langues se lisant de droite à gauche.) Pour plus d’informations sur la définition des [groupes et des contrôles de l’onglet, voir define the Groups and Controls](#define-the-groups-and-controls-that-appear-on-the-tab) .
- Si plusieurs compléments disposent d’un onglet contextuel visible dans un contexte spécifique, ils apparaissent dans l’ordre dans lequel les compléments ont été lancés.
- Les onglets *contextuels* personnalisés, contrairement aux onglets principaux personnalisés, ne sont pas ajoutés de façon permanente au ruban de l’application Office. Ils sont présents uniquement dans les documents Office sur lesquels votre complément est en cours d’exécution.

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>Étapes principales pour l’inclusion d’un onglet contextuel dans un complément

Voici les principales étapes à suivre pour inclure un onglet contextuel personnalisé dans un complément :

1. Configurez le complément pour qu’il utilise un runtime partagé.
1. Définissez l’onglet, ainsi que les groupes et les contrôles qui s’y trouvent.
1. Enregistrer l’onglet contextuel avec Office.
1. Spécifier les circonstances dans lesquelles l’onglet sera visible.

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurer le complément pour utiliser un runtime partagé

L’ajout d’onglets contextuels personnalisés nécessite que votre complément utilise le runtime partagé. Pour plus d’informations, consultez [la rubrique Configure an Add-in use a Shared Runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>Définir les groupes et les contrôles qui apparaissent sur l’onglet

Contrairement aux onglets principaux personnalisés, qui sont définis avec XML dans le manifeste, les onglets contextuels personnalisés sont définis au moment de l’exécution avec un blob JSON. Votre code analyse le BLOB dans un objet JavaScript, puis passe l’objet à la méthode [Office. Ribbon. requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) . Les onglets contextuels personnalisés sont présents uniquement dans les documents sur lesquels votre complément est en cours d’exécution. Cette fonction est différente des onglets principaux personnalisés qui sont ajoutés au ruban de l’application Office lorsque le complément est installé et reste présent lors de l’ouverture d’un autre document. De plus, la `requestCreateControls` méthode ne peut être exécutée qu’une seule fois dans une session de votre complément. Si elle est encore appelée, une erreur est générée.

> [!NOTE]
> La structure des propriétés et sous-propriétés du BLOB JSON (et les noms de clés) est à peu près parallèle à la structure de l’élément [CustomTab](../reference/manifest/customtab.md) et ses éléments descendants dans le fichier manifeste XML.

Nous allons construire un exemple d’objet de BLOB JSON d’onglets contextuels pas à pas. (Le schéma complet pour l’onglet contextuel JSON se trouve [ surdynamic-ribbon.schema.js](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json). Ce lien peut ne pas fonctionner dans la période d’aperçu anticipé des onglets contextuels. Si le lien ne fonctionne pas, vous pouvez trouver le dernier brouillon du schéma [sur brouillon dynamic-ribbon.schema.jssur](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json).) Si vous travaillez dans Visual Studio code, vous pouvez utiliser ce fichier pour obtenir IntelliSense et valider votre JSON. Pour plus d’informations, consultez la rubrique [Editing JSON with Visual Studio code-JSON schemas and Settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).


1. Commencez par créer une chaîne JSON avec deux propriétés de tableau nommées `actions` et `tabs` . Le `actions` tableau est une spécification de toutes les fonctions qui peuvent être exécutées par des contrôles dans l’onglet contextuel. Le `tabs` tableau définit un ou plusieurs onglets contextuels, *jusqu’à un maximum de 10*.

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. Cet exemple simple d’onglet contextuel ne comporte qu’un seul bouton et, par conséquent, une seule action. Ajoutez ce qui suit en tant que membre unique du `actions` tableau. À propos de ce balisage, notez les éléments suivants :

    - Les `id` `type` Propriétés et sont obligatoires.
    - La valeur de `type` peut être « ExecuteFunction » ou « ShowTaskpane ».
    - La `functionName` propriété est utilisée uniquement lorsque la valeur de `type` est `ExecuteFunction` . Il s’agit du nom d’une fonction définie dans FunctionFile. Pour plus d’informations sur les FunctionFile, consultez la rubrique [concepts de base pour les commandes de complément](add-in-commands.md).
    - Dans une étape ultérieure, vous mapperez cette action sur un bouton de l’onglet contextuel.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. Ajoutez ce qui suit en tant que membre unique du `tabs` tableau. À propos de ce balisage, notez les éléments suivants :

    - La propriété `id` est requise. Utilisez un bref ID descriptif qui est unique parmi tous les onglets contextuels de votre complément.
    - La propriété `label` est requise. Il s’agit d’une chaîne conviviale qui servira d’étiquette de l’onglet contextuel.
    - La propriété `groups` est requise. Il définit les groupes de contrôles qui s’affichent sous l’onglet. Il doit comporter au moins un membre *et pas plus de 20*. (Il existe également des limites quant au nombre de contrôles que vous pouvez avoir sur un onglet contextuel personnalisé et qui contraignent également le nombre de groupes dont vous disposez. Pour plus d’informations, reportez-vous à l’étape suivante.)

    > [!NOTE]
    > L’objet Tab peut également avoir une `visible` propriété facultative qui spécifie si l’onglet est visible immédiatement au démarrage du complément. Étant donné que les onglets contextuels sont normalement masqués jusqu’à ce qu’un événement utilisateur déclenche sa visibilité (par exemple, l’utilisateur sélectionnant une entité d’un certain type dans le document), la `visible` propriété est définie par défaut `false` lorsqu’il n’est pas présent. Dans une section ultérieure, nous montrons comment définir la propriété sur `true` en réponse à un événement.

    ```json
    {
      "id": "CtxTab1",
      "label": "Data",
      "groups": [

      ]
    }
    ```

1. Dans l’exemple simple en cours, l’onglet contextuel ne comporte qu’un seul groupe. Ajoutez ce qui suit en tant que membre unique du `groups` tableau. À propos de ce balisage, notez les éléments suivants :

    - Toutes les propriétés sont requises.
    - La `id` propriété doit être unique parmi tous les groupes de l’onglet. Utilisez un bref ID descriptif.
    - `label`Est une chaîne conviviale qui sert d’étiquette au groupe.
    - La `icon` valeur de la propriété est un tableau d’objets qui spécifient les icônes du groupe sur le ruban en fonction de la taille du ruban et de la fenêtre de l’application Office.
    - La `controls` valeur de la propriété est un tableau d’objets qui spécifient les boutons et d’autres contrôles dans le groupe. Il doit y avoir au moins un et *six pour un groupe*.

    > [!IMPORTANT]
    > *Le nombre total de contrôles sur l’onglet entier ne peut pas être supérieur à 20.* Par exemple, vous pouvez avoir 3 groupes avec 6 contrôles chacun, et un quatrième groupe avec 2 contrôles, mais vous ne pouvez pas avoir 4 groupes avec 6 contrôles chacun.  

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

1. Chaque groupe doit avoir une icône d’au moins deux tailles, 32x32 PX et 80x80 px. Vous pouvez également avoir des icônes de taille 16x16, 20x20, 24x24, 40x40, 48 x 48 et 64 x 64. Office décide de l’icône à utiliser en fonction de la taille du ruban et de la fenêtre de l’application Office. Ajoutez les objets suivants au tableau d’icônes. (Si la taille de la fenêtre et du ruban est suffisante pour qu’au moins un des *contrôles* du groupe s’affiche, aucune icône de groupe n’apparaît. Pour obtenir un exemple, Regardez le groupe **styles** sur le ruban Word lorsque vous réduisez et développez la fenêtre Word.) À propos de ce balisage, notez les éléments suivants :

    - Les deux propriétés sont requises.
    - L' `size` unité de mesure de la propriété est exprimée en pixels. Les icônes sont toujours carrées, de sorte que le nombre est à la fois la hauteur et la largeur.
    - La `sourceLocation` propriété spécifie l’URL complète de l’icône.

    > [!IMPORTANT]
    > Tout comme vous devez généralement modifier les URL dans le manifeste du complément lorsque vous passez du développement à la production (par exemple, en modifiant le domaine de localhost vers contoso.com), vous devez également modifier les URL dans vos onglets contextuels JSON.

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

1. Dans notre exemple simple en cours, le groupe ne possède qu’un seul bouton. Ajoutez l’objet suivant en tant que membre unique du `controls` tableau. À propos de ce balisage, notez les éléments suivants :

    - Toutes les propriétés, à l’exception `enabled` de, sont obligatoires.
    - `type` Spécifie le type de contrôle. Les valeurs peuvent être « Button », « menu » ou « MobileButton ».
    - `id` peut contenir jusqu’à 125 caractères. 
    - `actionId` doit être l’ID d’une action définie dans le `actions` tableau. (Voir l’étape 1 de cette section.)
    - `label` est une chaîne conviviale à utiliser comme étiquette du bouton.
    - `superTip` représente une forme enrichie d’info-bulle. Les `title` Propriétés et `description` sont toutes deux requises.
    - `icon` spécifie les icônes du bouton. Les remarques précédentes sur l’icône de groupe s’appliquent également ici.
    - `enabled` (facultatif) indique si le bouton est activé lorsque l’onglet contextuel s’affiche. La valeur par défaut est `true` . 

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
 
Voici l’exemple complet de l’objet BLOB JSON :

```json
'{
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
      "label": "Data",
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
}'
```

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>Enregistrer l’onglet contextuel avec Office avec requestCreateControls

L’onglet contextuel est inscrit avec Office en appelant la méthode [Office. Ribbon. requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) . Cette opération s’effectue généralement dans la fonction qui est affectée à `Office.initialize` ou avec la `Office.onReady` méthode. Pour plus d’informations sur ces méthodes et l’initialisation du complément, reportez-vous à la rubrique [initialiser votre complément Office](../develop/initialize-add-in.md). Toutefois, vous pouvez appeler la méthode à tout moment après l’initialisation.

> [!IMPORTANT]
> La `requestCreateControls` méthode ne peut être appelée qu’une seule fois dans une session donnée d’un complément. Une erreur est générée si elle est encore appelée.

Voici un exemple. Notez que la chaîne JSON doit être convertie en objet JavaScript avec la `JSON.parse` méthode avant de pouvoir être transmise à une fonction JavaScript.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>Spécifier les contextes lorsque l’onglet est visible avec requestUpdate

En règle générale, un onglet contextuel personnalisé doit s’afficher lorsqu’un événement initié par l’utilisateur modifie le contexte du complément. Imaginez un scénario dans lequel l’onglet doit être visible lorsque, et uniquement lorsqu’un graphique (de la feuille de calcul par défaut d’un classeur Excel) est activé.

Commencez par attribuer des gestionnaires. Cette opération est généralement exécutée dans la `Office.onReady` méthode, comme dans l’exemple suivant, qui affecte des gestionnaires (créés à une étape ultérieure) aux `onActivated` `onDeactivated` événements et de tous les graphiques de la feuille de calcul.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string.
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

Ensuite, définissez les gestionnaires. Voici un exemple simple d’un `showDataTab` , mais voir la [gestion des erreurs](#error-handling) plus loin dans cet article pour obtenir une version plus robuste de la fonction. Tenez compte du code suivant :

- Office effectue un contrôle lorsqu’il met à jour l’état du ruban. La méthode  [Office. Ribbon. requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) met en file d’attente une requête à mettre à jour. La méthode permet de résoudre l’objet dès qu' `Promise` il a mis en file d’attente la demande, et non lors de la mise à jour du ruban.
- Le paramètre de la `requestUpdate` méthode est un objet [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) qui (1) spécifie l’onglet par son ID *exactement comme spécifié dans JSON* et (2) indique la visibilité de l’onglet.
- Si vous avez plusieurs onglets contextuels personnalisés qui doivent être visibles dans le même contexte, il vous suffit d’ajouter des objets Tab supplémentaires au `tabs` tableau.

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

Le gestionnaire permettant de masquer l’onglet est quasiment identique, à la différence qu’il rétablit la `visible` propriété sur `false` .

La bibliothèque JavaScript Office fournit également plusieurs interfaces (types) pour faciliter la création de l' `RibbonUpdateData` objet. Voici la `showDataTab` fonction dans la machine à écrire et elle utilise ces types.

```typescript
const showDataTab = async () => {
    const myContextualTab: Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Activer/désactiver la visibilité de l’onglet et l’état activé d’un bouton en même temps

La `requestUpdate` méthode est également utilisée pour faire basculer l’état activé ou désactivé d’un bouton personnalisé dans un onglet contextuel personnalisé ou un onglet de base personnalisé. Pour plus d’informations à ce sujet, consultez la rubrique [activer et désactiver des commandes de complément](disable-add-in-commands.md). Il peut y avoir des scénarios dans lesquels vous souhaitez modifier à la fois la visibilité d’un onglet et l’état activé d’un bouton en même temps. Vous pouvez effectuer cette opération à l’aide d’un seul appel de `requestUpdate` . Voici un exemple dans lequel un bouton d’un onglet principal est activé en même temps qu’un onglet contextuel est rendu visible.

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
                controls: [
                {
                    id: "MyButton",
                    enabled: true
                }
            ]}
        ]});
}
```

Dans l’exemple suivant, le bouton activé se trouve sur le même onglet contextuel qui est rendu visible.

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                controls: [
                    {
                        id: "MyButton",
                        enabled: true
                    }
                ]
            }
        ]});
}
```

## <a name="error-handling"></a>Gestion des erreurs

Dans certains scénarios, Office ne peut pas mettre à jour le ruban et renvoie une erreur. Par exemple, si le complément est mis à niveau et que le complément mis à niveau dispose d'un autre groupe de commandes de complément personnalisé, l’application Office doit être fermée et ouverte de nouveau. La méthode `requestUpdate` renvoie l'erreur `HostRestartNeeded` jusqu'à ce que cela soit effectué. Voici comment vous pouvez gérer cette erreur. Dans ce cas, la méthode `reportError` affiche l’erreur à l’utilisateur.

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
