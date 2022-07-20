---
title: Commandes Activé et Désactivé pour les compléments
description: Découvrez la modification de l'état Activé ou Désactivé des boutons de rubans et des éléments de menu personnalisés dans votre complément web Office.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: c16d82b1c6dacc3e46f2a5170c4b43d1be899191
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889148"
---
# <a name="enable-and-disable-add-in-commands"></a>Commandes Activé et Désactivé pour les compléments

Lorsque seulement quelques fonctionnalités de votre complément doivent être disponibles dans certains contextes, vous pouvez activer ou désactiver vos commandes de complément personnalisées par programme. Par exemple, une fonction qui modifie l’en-tête d’un tableau doit être uniquement activée lorsque le curseur se trouve dans un tableau.

Vous pouvez également spécifier si la commande est activée ou désactivée lorsque l’application cliente Office s’ouvre.

> [!NOTE]
> Cet article suppose que vous connaissez la documentation décrite ci-après. Étudiez-la si vous n’avez pas récemment utilisé les commandes de complément (éléments de menu et boutons de ruban personnalisés).
>
> - [Concepts basiques pour les commandes de complément](add-in-commands.md)

## <a name="office-application-and-platform-support-only"></a>Prise en charge des applications et plateformes Office uniquement

Les API décrites dans cet article sont disponibles uniquement dans Excel, PowerPoint et Word.

### <a name="test-for-platform-support-with-requirement-sets"></a>Effectuez un test pour la prise en charge des plateformes avec les ensembles de conditions requises

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent des ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification d’exécution pour déterminer si une combinaison d’applications et de plateformes Office prend en charge les API dont un complément a besoin. Pour plus d’informations, consultez [les versions d’Office et les ensembles de conditions requises](../develop/office-versions-and-requirement-sets.md).

Les API enable/disable appartiennent à l’ensemble de conditions requises [RibbonApi 1.1](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets) .

> [!NOTE]
> L’ensemble de conditions requises **RibbonApi 1.1** n’étant pas encore pris en charge dans le manifeste, vous ne pouvez pas le spécifier dans la section du **\<Requirements\>** manifeste. Pour tester la prise en charge, votre code doit appeler `Office.context.requirements.isSetSupported('RibbonApi', '1.1')`. Si, *et seulement si*, cet appel est retourné `true`, votre code peut appeler les API enable/disable. Si l’appel de `isSetSupported` retours `false`est retourné, toutes les commandes de complément personnalisées sont activées tout le temps. Vous devez concevoir votre complément de production et toutes les instructions dans l’application pour prendre en compte son fonctionnement lorsque l’ensemble de conditions requises **RibbonApi 1.1** n’est pas pris en charge. Pour plus d’informations et des exemples d’utilisation `isSetSupported`, consultez [Spécifier les applications Office et les exigences d’API](../develop/specify-office-hosts-and-api-requirements.md), en particulier les [vérifications runtime pour la prise en charge de la méthode et de l’ensemble de conditions requises](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). (La section [Spécifier les versions et plateformes Office pouvant héberger votre complément](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in) de cet article ne s’applique pas au ruban 1.1.)

## <a name="shared-runtime-required"></a>Runtime partagé requis

Les API et balisages de manifeste décrits dans cet article exigent que le manifeste du complément spécifie la nécessité d’utiliser un runtime partagé. Pour ce faire, procédez comme suit.

1. Dans l'élément [Runtimes du manifeste](/javascript/api/manifest/runtimes), ajoutez l’élément enfant suivant : `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`. (S’il n’existe pas encore d’élément **\<Runtimes\>** dans le manifeste, créez-le en tant que premier enfant sous l’élément **\<Host\>** de la **\<VersionOverrides\>** section.)
2. Dans la section [Resources.Urls](/javascript/api/manifest/resources) du manifeste, ajoutez l’élément enfant suivant : `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, où `{MyDomain}` est le domaine du complément et `{path-to-start-page}` le chemin d’accès de la page de démarrage du complément. par exemple : `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.
3. Selon que votre complément contient un volet Office, un fichier de fonction ou une fonction personnalisée Excel, vous devez effectuer une ou plusieurs des trois étapes suivantes.

    - Si le complément contient un volet Office, définissez l’attribut `resid` de [l’action](/javascript/api/manifest/action).[ Élément SourceLocation](/javascript/api/manifest/sourcelocation) à exactement la même chaîne que vous avez utilisée pour l’élément `resid` à l’étape **\<Runtime\>** 1 ; par exemple, `Contoso.SharedRuntime.Url`. Le fichier doit ressembler à ceci : `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - Si le complément contient une fonction personnalisée Excel, définissez l’attribut `resid` de la [page](/javascript/api/manifest/page).[ Élément SourceLocation](/javascript/api/manifest/sourcelocation) exactement la même chaîne que vous avez utilisée pour l’élément `resid` à l’étape **\<Runtime\>** 1 ; par exemple, `Contoso.SharedRuntime.Url`. Le fichier doit ressembler à ceci : `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - Si le complément contient un fichier de fonction, définissez l’attribut `resid` de l’élément [FunctionFile](/javascript/api/manifest/functionfile) sur exactement la même chaîne que celle que vous avez utilisée pour l’élément `resid` à l’étape **\<Runtime\>** 1 ; par exemple, `Contoso.SharedRuntime.Url`. Le fichier doit ressembler à ceci : `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.

## <a name="set-the-default-state-to-disabled"></a>Configurer l'état par défaut sur désactivé

Les commandes de complément sont activées par défaut au démarrage de l’application Office. Si vous souhaitez qu’un bouton ou un élément de menu personnalisé soit désactivé au démarrage de l’application Office, vous devez le spécifier dans le manifeste. Il vous suffit d’ajouter un élément [activé](/javascript/api/manifest/enabled) (avec la valeur `false`) juste *au-dessous* (non à l’intérieur) de l'élément [Action](/javascript/api/manifest/action) dans la déclaration du contrôle. L’exemple suivant montre la structure de base.

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
                <Control ... id="Contoso.MyButton3">
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a>Modifier l’état par programme

Les principales étapes pour modifier l’état activé d’une commande de complément sont les suivantes :

1. Créez un objet [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) qui (1) spécifie la commande, ainsi que son groupe parent et son onglet, par leurs ID déclarés dans le manifeste ; et (2) spécifie l’état activé ou désactivé de la commande.
2. Transmettez l’objet **RibbonUpdaterData** à la méthode [Office.ribbon.requestUpdate ()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)).

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
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

Vous pouvez `await` appeler **requestUpdate()** si la fonction parente est asynchrone, mais notez que l’application Office contrôle quand elle met à jour l’état du ruban. La méthode **requestUpdate()** met en file d’attente une demande de mise à jour. La méthode résout l’objet de promesse dès qu’il a mis en file d’attente la requête, et non quand le ruban est réellement mis à jour.

## <a name="change-the-state-in-response-to-an-event"></a>Modifier l’état en réponse à un événement

Un scénario courant est celui lors duquel l’état du ruban peut être modifié lorsqu’un événement initié par l’utilisateur modifie le contexte du complément.

Imaginez un scénario dans lequel un bouton doit être activé lorsque, et seulement lorsqu'un graphique est activé. La première étape consiste à définir l'élément [Activé](/javascript/api/manifest/enabled) pour le bouton dans le manifeste `false`. Voir l'exemple ci-dessus.

Deuxièmement, assignez des gestionnaires. Cette procédure est généralement effectuée dans la méthode **Office.onReady** comme illustré dans l’exemple suivant qui assigne des gestionnaires (créés dans une étape ultérieure) aux évènements **onActivated** et **onDeactivated** de tous les graphiques de la feuille de calcul.

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        const charts = context.workbook.worksheets
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
    const button = {
                  id: "ChartFormatButton", 
                  enabled: true
                 };
    const parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    const parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    const ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

Quatrièmement, définissez le gestionnaire `disableChartFormat`. Il est identique à `enableChartFormat`, sauf que la propriété **activé** de l’objet bouton a la valeur `false`.

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Activer simultanément la visibilité de l’onglet et l’état activé d’un bouton

La méthode **requestUpdate** est également utilisée pour désactiver la visibilité d’un onglet contextuel personnalisé. Pour plus d’informations sur ce code et l’exemple de code, consultez [Créer des onglets contextuels personnalisés dans les compléments Office](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time).

## <a name="best-practice-test-for-control-status-errors"></a>Pratiques recommandées : test pour les erreurs de contrôle d'état

Le ruban ne se redessine pas, dans certains cas, une fois que `requestUpdate` est appelé, de sorte que l’état du contrôle cliquable ne change pas. Il est pour cette raison recommandé de suivre l'état des contrôles du complément. Le complément doit respecter les règles suivantes.

1. Lorsque `requestUpdate` est appelé, le code doit enregistrer l’état prévu des boutons et éléments de menu personnalisés.
2. Lorsque l’utilisateur clique sur un contrôle personnalisé, le premier code dans le gestionnaire doit vérifier si le bouton aurait dû être cliquable. Dans la négative, le code doit signaler une erreur ou consigner une erreur et réessayer de définir les boutons de l'état prévu.

L’exemple suivant présente une fonction qui désactive un bouton et enregistre l’état du bouton. Veuillez noter que `chartFormatButtonEnabled` est une variable Boolean globale qui est initialisée sur la même valeur que l'élément [Activé](/javascript/api/manifest/enabled) pour le bouton dans le manifeste.

```javascript
function disableChartFormat() {
    const button = {
                  id: "ChartFormatButton", 
                  enabled: false
                 };
    const parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    const parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    const ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);

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
        const button = {
                      id: "ChartFormatButton", 
                      enabled: false
                     };
        const parentGroup = {
                           id: "MyGroup",
                           controls: [button]
                          };
        const parentTab = {
                         id: "CustomChartTab", 
                         groups: [parentGroup]
                        };
        const ribbonUpdater = {tabs: [parentTab]};
        Office.ribbon.requestUpdate(ribbonUpdater);

        chartFormatButtonEnabled = false;
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, close the Office application, and restart it.");
        }
    }
}
```
