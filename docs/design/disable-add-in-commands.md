---
title: Commandes Activé et Désactivé pour les compléments
description: Découvrez la modification de l'état Activé ou Désactivé des boutons de rubans et des éléments de menu personnalisés dans votre complément web Office.
ms.date: 01/12/2021
localization_priority: Normal
ms.openlocfilehash: 798dd723e0388becdd3419c5af87ceb360d32a41
ms.sourcegitcommit: 6a378d2a3679757c5014808ae9da8ababbfe8b16
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/15/2021
ms.locfileid: "49870629"
---
# <a name="enable-and-disable-add-in-commands"></a>Commandes Activé et Désactivé pour les compléments

Lorsque seulement quelques fonctionnalités de votre complément doivent être disponibles dans certains contextes, vous pouvez activer ou désactiver vos commandes de complément personnalisées par programme. Par exemple, une fonction qui modifie l’en-tête d’un tableau doit être uniquement activée lorsque le curseur se trouve dans un tableau.

Vous pouvez également spécifier si la commande est activée ou désactivée à l’ouverture de l’application cliente Office.

> [!NOTE]
> Cet article suppose que vous connaissez la documentation décrite ci-après. Étudiez-la si vous n’avez pas récemment utilisé les commandes de complément (éléments de menu et boutons de ruban personnalisés).
>
> - [Concepts basiques pour les commandes de complément](add-in-commands.md)

## <a name="office-application-and-platform-support-only"></a>Prise en charge des applications et des plateformes Office uniquement

Les API décrites dans cet article sont disponibles uniquement dans Excel, et uniquement dans Office sur Windows, Office sur Mac et Office sur le web.

### <a name="test-for-platform-support-with-requirement-sets"></a>Effectuez un test pour la prise en charge des plateformes avec les ensembles de conditions requises

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les applications Office utilisent des ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’runtime pour déterminer si une combinaison d’application Office et de plateforme prend en charge les API requises par un application. Pour plus d’informations, voir [les versions d’Office et les ensembles de conditions requises.](../develop/office-versions-and-requirement-sets.md)

Les API d’activer/désactiver appartiennent à [l’ensemble de conditions requises RibbonApi 1.1.](../reference/requirement-sets/ribbon-api-requirement-sets.md)

> [!NOTE]
> L’ensemble de conditions requises **RibbonApi 1.1** n’est pas encore pris en charge dans le manifeste, vous ne pouvez donc pas le spécifier dans la section du `<Requirements>` manifeste. Pour tester la prise en charge, votre code doit appeler `Office.context.requirements.isSetSupported('RibbonApi', '1.1')` . Si, *et uniquement si*, cet appel renvoie, votre code peut appeler les API `true` d’activer/désactiver. Si l’appel de renvoie, toutes les commandes de modules sont `isSetSupported` `false` activées en temps réel. Vous devez concevoir votre application de production, ainsi que les instructions dans l’application, pour prendre en compte son fonctionnement lorsque l’ensemble de conditions requises **RibbonApi 1.1** n’est pas pris en charge. Pour plus d’informations et des exemples d’utilisation, voir Spécifier les applications Office et les conditions requises de l’API, en particulier utiliser les vérifications à l’runtime `isSetSupported` dans votre code [JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code) [](../develop/specify-office-hosts-and-api-requirements.md). (La section [Définir l’élément Requirements dans le manifeste](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) de cet article ne s’applique pas au Ruban 1.1.)

## <a name="shared-runtime-required"></a>Runtime partagé requis

Les API et balisages de manifeste décrits dans cet article exigent que le manifeste du complément spécifie la nécessité d’utiliser un runtime partagé. Pour ce faire, procédez comme suit.

1. Dans l'élément [Runtimes du manifeste](../reference/manifest/runtimes.md), ajoutez l’élément enfant suivant : `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`. (s’il n’y a pas encore d’élément `<Runtimes>` dans le manifeste, créez-le en tant que premier enfant sous l’élément `<Host>` dans la section `VersionOverrides`.)
2. Dans la section [Resources.Urls](../reference/manifest/resources.md) du manifeste, ajoutez l’élément enfant suivant : `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, où `{MyDomain}` est le domaine du complément et `{path-to-start-page}` le chemin d’accès de la page de démarrage du complément. par exemple : `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.
3. En fonction du contenu de votre complément, à savoir un volet des tâches, un fichier de fonctions ou une fonction Excel personnalisée, vous devez effectuer au moins une des trois étapes suivantes :

    - Si le complément contient un volet Office, définissez l'attribut `resid` de l’élément [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) sur la même chaîne que celle que vous avez utilisée pour le `resid` de l’élément `<Runtime>` à l’étape 1 ; par exemple, `Contoso.SharedRuntime.Url`. Le fichier doit ressembler à ceci : `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - Si le complément contient une fonction personnalisée Excel, définissez l'attribut `resid` de l’élément [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) sur la même chaîne que celle que vous avez utilisée pour le `resid` de l’élément `<Runtime>` à l’étape 1 ; par exemple, `Contoso.SharedRuntime.Url`. Le fichier doit ressembler à ceci : `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - Si le complément contient un fichier de fonctions, définissez l'attribut `resid` de l’élément [FunctionFile](../reference/manifest/functionfile.md) sur la même chaîne que celle que vous avez utilisée pour le `resid` de l’élément `<Runtime>` à l’étape 1 ; par exemple, `Contoso.SharedRuntime.Url`. Le fichier doit ressembler à ceci : `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.

## <a name="set-the-default-state-to-disabled"></a>Configurer l'état par défaut sur désactivé

Les commandes de complément sont activées par défaut au démarrage de l’application Office. Si vous souhaitez qu’un bouton ou un élément de menu personnalisé soit désactivé au démarrage de l’application Office, vous devez le spécifier dans le manifeste. Il vous suffit d’ajouter un élément [activé](../reference/manifest/enabled.md) (avec la valeur `false`) juste *au-dessous* (non à l’intérieur) de l'élément [Action](../reference/manifest/action.md) dans la déclaration du contrôle. La structure basique est la suivante :

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
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a>Modifier l’état par programme

Les principales étapes pour modifier l’état activé d’une commande de complément sont les suivantes :

1. Créez [un objet RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) qui (1) spécifie la commande, ainsi que son groupe parent et son onglet, par leurs ID déclarés dans le manifeste ; et (2) spécifie l’état activé ou désactivé de la commande.
2. Transmettez l’objet **RibbonUpdaterData** à la méthode [Office.ribbon.requestUpdate ()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-).

Voici un exemple simple. Notez que « MyButton », « OfficeAddinTab1 » et « CustomGroup111 » sont copiés à partir du manifeste.

```javascript
function enableButton() {
    Office.ribbon.requestUpdate({
        tabs: [
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
            }
        ]
    });
}
```

Nous proposons également plusieurs interfaces (types) pour faciliter la construction de l’objet **RibbonUpdateData**. L’exemple suivant est l’équivalent de TypeScript et il utilise ces types.

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentGroup: Group = {id: "CustomGroup111", controls: [button]};
    const parentTab: Tab = {id: "OfficeAddinTab1", groups: [parentGroup]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

Office effectue un contrôle lorsqu’il met à jour l’état du ruban. La méthode **requestUpdate()** met en file d’attente une demande de mise à jour. La méthode permet de résoudre l’objet Promesse dès la mise en file d’attente de la demande, et non lors la prochaine mise à jour du ruban.

## <a name="change-the-state-in-response-to-an-event"></a>Modifier l’état en réponse à un événement

Un scénario courant est celui lors duquel l’état du ruban peut être modifié lorsqu’un événement initié par l’utilisateur modifie le contexte du complément.

Imaginez un scénario dans lequel un bouton doit être activé lorsque, et seulement lorsqu'un graphique est activé. La première étape consiste à définir l'élément [Activé](../reference/manifest/enabled.md) pour le bouton dans le manifeste `false`. Voir l'exemple ci-dessus.

Deuxièmement, assignez des gestionnaires. Cette procédure est généralement effectuée dans la méthode **Office.onReady** comme illustré dans l’exemple suivant qui assigne des gestionnaires (créés dans une étape ultérieure) aux évènements **onActivated** et **onDeactivated** de tous les graphiques de la feuille de calcul.

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(enableChartFormat);
        charts.onDeactivated.add(disableChartFormat);
        return context.sync();
    });
});
```

Troisièmement, définissez le gestionnaire `enableChartFormat`. Voici un exemple simple, mais consultez les [Pratiques recommandées : test pour les erreurs de contrôle d’état](#best-practice-test-for-control-status-errors) ci-dessous pour modifier l’état d’un contrôle de façon plus efficace.

```javascript
function enableChartFormat() {
    var button = {
                  id: "ChartFormatButton", 
                  enabled: true
                 };
    var parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    var parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

Quatrièmement, définissez le gestionnaire `disableChartFormat`. Il est identique à `enableChartFormat`, sauf que la propriété **activé** de l’objet bouton a la valeur `false`.

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Activer la visibilité de l’onglet et l’état activé d’un bouton en même temps

La **méthode requestUpdate** est également utilisée pour faire bascule la visibilité d’un onglet contextuel personnalisé. Pour plus d’informations sur ce code et un exemple de code, voir [Créer des onglets contextuels personnalisés dans les add-ins Office.](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time)

## <a name="best-practice-test-for-control-status-errors"></a>Pratiques recommandées : test pour les erreurs de contrôle d'état

Le ruban ne se redessine pas, dans certains cas, une fois que `requestUpdate` est appelé, de sorte que l’état du contrôle cliquable ne change pas. Il est pour cette raison recommandé de suivre l'état des contrôles du complément. Le complément doit respecter les règles suivantes :

1. Lorsque `requestUpdate` est appelé, le code doit enregistrer l’état prévu des boutons et éléments de menu personnalisés.
2. Lorsque l’utilisateur clique sur un contrôle personnalisé, le premier code dans le gestionnaire doit vérifier si le bouton aurait dû être cliquable. Dans la négative, le code doit signaler une erreur ou consigner une erreur et réessayer de définir les boutons de l'état prévu.

L’exemple suivant présente une fonction qui désactive un bouton et enregistre l’état du bouton. Veuillez noter que `chartFormatButtonEnabled` est une variable Boolean globale qui est initialisée sur la même valeur que l'élément [Activé](../reference/manifest/enabled.md) pour le bouton dans le manifeste.

```javascript
function disableChartFormat() {
    var button = {
                  id: "ChartFormatButton", 
                  enabled: false
                 };
    var parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    var parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

L’exemple suivant présente la façon dont le gestionnaire du bouton vérifie l’état d’un bouton incorrect. Veuillez noter que `reportError` est une fonction qui affiche ou consigne une erreur.

```javascript
function chartFormatButtonHandler() {
    if (chartFormatButtonEnabled) {

        // Do work here

    } else {
        // Report the error and try again to disable.
        reportError("That action is not possible at this time.");
        disableChartFormat();
    }
}
```

## <a name="error-handling"></a>Gestion des erreurs

Dans certains scénarios, Office ne peut pas mettre à jour le ruban et renvoie une erreur. Par exemple, si le complément est mis à niveau et que le complément mis à niveau dispose d'un autre groupe de commandes de complément personnalisé, l’application Office doit être fermée et ouverte de nouveau. La méthode `requestUpdate` renvoie l'erreur `HostRestartNeeded` jusqu'à ce que cela soit effectué. Voici comment vous pouvez gérer cette erreur. Dans ce cas, la méthode `reportError` affiche l’erreur à l’utilisateur.

```javascript
function disableChartFormat() {
    try {
        var button = {
                      id: "ChartFormatButton", 
                      enabled: false
                     };
        var parentGroup = {
                           id: "MyGroup",
                           controls: [button]
                          };
        var parentTab = {
                         id: "CustomChartTab", 
                         groups: [parentGroup]
                        };
        var ribbonUpdater = {tabs: [parentTab]};
        await Office.ribbon.requestUpdate(ribbonUpdater);

        chartFormatButtonEnabled = false;
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, close the Office application, and restart it.");
        }
    }
}
```
