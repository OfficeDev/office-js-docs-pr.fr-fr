---
title: Consigner les notes de rendez-vous dans une application externe dans les compléments mobiles Outlook
description: Découvrez comment configurer un complément mobile Outlook pour consigner les notes de rendez-vous et d’autres détails dans une application externe.
ms.topic: article
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: a980b68c603154c42112f525ec6285b740ce38a5
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607582"
---
# <a name="log-appointment-notes-to-an-external-application-in-outlook-mobile-add-ins"></a>Consigner les notes de rendez-vous dans une application externe dans les compléments mobiles Outlook

L’enregistrement de vos notes de rendez-vous et d’autres détails dans une application CRM (Customer Relationship Management) ou une application de prise de notes peut vous aider à effectuer le suivi des réunions auxquelles vous avez participé.

Dans cet article, vous allez apprendre à configurer votre complément mobile Outlook pour permettre aux utilisateurs de consigner des notes et d’autres détails sur leurs rendez-vous dans votre application CRM ou de prise de notes. Tout au long de cet article, nous allons utiliser un fournisseur de services CRM fictif nommé « Contoso ».

> [!IMPORTANT]
> Cette fonctionnalité est uniquement prise en charge sur Android avec un abonnement Microsoft 365.

## <a name="set-up-your-environment"></a>Configuration de votre environnement

Terminez le [démarrage rapide d’Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) pour créer un projet de complément avec le générateur Yeoman pour les compléments Office.

## <a name="capture-and-view-appointment-notes"></a>Capturer et afficher les notes de rendez-vous

Vous pouvez choisir d’implémenter une commande de fonction ou un volet Office. Pour mettre à jour votre complément, sélectionnez l’onglet pour la commande de fonction ou le volet Office, puis suivez les instructions.

# <a name="function-command"></a>[Commande de fonction](#tab/noui)

Cette option permet à un utilisateur de journaliser et d’afficher ses notes et d’autres détails sur ses rendez-vous lorsqu’il sélectionne une commande de fonction dans le ruban.

### <a name="configure-the-manifest"></a>Configurer le manifeste

Pour permettre aux utilisateurs de consigner les notes de rendez-vous avec votre complément, vous devez configurer le [point d’extension MobileLogEventAppointmentAttendee](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee) dans le manifeste sous l’élément `MobileFormFactor`parent. D’autres facteurs de forme ne sont pas pris en charge.

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

1. Dans votre éditeur de code, ouvrez le projet de démarrage rapide.

1. Ouvrez le fichier **manifest.xml** situé à la racine de votre projet.

1. Sélectionnez l’intégralité `<VersionOverrides>` du nœud (y compris les balises d’ouverture et de fermeture) et remplacez-le par le code XML suivant. Veillez à remplacer toutes les références à **Contoso** par les informations de votre entreprise.

    ```xml
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
      <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
        <Description resid="residDescription"></Description>
        <Requirements>
          <bt:Sets>
            <bt:Set Name="Mailbox" MinVersion="1.3"/>
          </bt:Sets>
        </Requirements>
        <Hosts>
          <Host xsi:type="MailHost">
            <DesktopFormFactor>
              <FunctionFile resid="residFunctionFile"/>
              <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
                <OfficeTab id="TabDefault">
                  <Group id="apptReadGroup">
                    <Label resid="residDescription"/>
                    <Control xsi:type="Button" id="apptReadOpenPaneButton">
                      <Label resid="residLabel"/>
                      <Supertip>
                        <Title resid="residLabel"/>
                        <Description resid="residTooltip"/>
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="icon-16"/>
                        <bt:Image size="32" resid="icon-32"/>
                        <bt:Image size="80" resid="icon-80"/>
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>logCRMEvent</FunctionName>
                      </Action>
                    </Control>
                  </Group>
                </OfficeTab>
              </ExtensionPoint>
            </DesktopFormFactor>
            <MobileFormFactor>
              <FunctionFile resid="residFunctionFile"/>
              <ExtensionPoint xsi:type="MobileLogEventAppointmentAttendee">
                <Control xsi:type="MobileButton" id="appointmentReadFunctionButton">
                  <Label resid="residLabel"/>
                  <Icon>
                    <bt:Image size="25" scale="1" resid="icon-16"/>
                    <bt:Image size="25" scale="2" resid="icon-16"/>
                    <bt:Image size="25" scale="3" resid="icon-16"/>
                    <bt:Image size="32" scale="1" resid="icon-32"/>
                    <bt:Image size="32" scale="2" resid="icon-32"/>
                    <bt:Image size="32" scale="3" resid="icon-32"/>
                    <bt:Image size="48" scale="1" resid="icon-48"/>
                    <bt:Image size="48" scale="2" resid="icon-48"/>
                    <bt:Image size="48" scale="3" resid="icon-48"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>logCRMEvent</FunctionName>
                  </Action>
                </Control>
              </ExtensionPoint>
            </MobileFormFactor>
          </Host>
        </Hosts>
        <Resources>
          <bt:Images>
            <bt:Image id="icon-16" DefaultValue="https://contoso.com/assets/icon-16.png"/>
            <bt:Image id="icon-32" DefaultValue="https://contoso.com/assets/icon-32.png"/>
            <bt:Image id="icon-48" DefaultValue="https://contoso.com/assets/icon-48.png"/>
            <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
          </bt:Images>
          <bt:Urls>
            <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
          </bt:Urls>
          <bt:ShortStrings>
            <bt:String id="residDescription" DefaultValue="Log appointment notes and other details to Contoso CRM."/>
            <bt:String id="residLabel" DefaultValue="Log to Contoso CRM"/>
          </bt:ShortStrings>
          <bt:LongStrings>
            <bt:String id="residTooltip" DefaultValue="Log notes to Contoso CRM for this appointment."/>
          </bt:LongStrings>
        </Resources>
      </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> Pour en savoir plus sur les manifestes pour les compléments Outlook, consultez [les manifestes de complément Outlook](manifests.md) et [la prise en charge des commandes de complément pour Outlook Mobile](add-mobile-support.md).

### <a name="capture-appointment-notes"></a>Capturer les notes de rendez-vous

Dans cette section, découvrez comment votre complément peut extraire les détails du rendez-vous lorsque l’utilisateur sélectionne le bouton **Journal** .

1. À partir du même projet de démarrage rapide, ouvrez le fichier **./src/commands/commands.js** dans votre éditeur de code.

1. Remplacez l’intégralité du contenu du fichier **commands.js** par le code JavaScript suivant.

    ```js
    var event;

    Office.initialize = function (reason) {
      // Add any initialization code here.
    };

    function logCRMEvent(appointmentEvent) {
      event = appointmentEvent;
      console.log(`Subject: ${Office.context.mailbox.item.subject}`);
      Office.context.mailbox.item.body.getAsync(
        "html",
        { asyncContext: "This is passed to the callback" },
        function callback(result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true });
          } else {
            console.error("Failed to get body.");
            event.completed({ allowEvent: false });
          }
        }
      );
    }

    // Register the function.
    Office.actions.associate("logCRMEvent", logCRMEvent);
    ```

Ensuite, mettez à jour le fichier **commands.html** pour référencer **commands.js**.

1. À partir du même projet de démarrage rapide, ouvrez le fichier **./src/commands/commands.html** dans votre éditeur de code.

1. Recherchez et remplacez `<script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>` les éléments suivants :

    ```html
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
    <script type="text/javascript" src="commands.js"></script>
    ```

### <a name="view-appointment-notes"></a>Afficher les notes de rendez-vous

L’étiquette du bouton **Journal** peut être activée pour afficher **l’affichage** en définissant la propriété personnalisée **Journalisée des événements** réservée à cet effet. Lorsque l’utilisateur sélectionne le bouton **Afficher** , il peut consulter ses notes journalisées pour ce rendez-vous.

Votre complément définit l’expérience d’affichage des journaux. Par exemple, vous pouvez afficher les notes de rendez-vous journalisées dans une boîte de dialogue lorsque l’utilisateur sélectionne le bouton **Afficher** . Pour plus d’informations sur l’utilisation des dialogues, [reportez-vous à l’API Utiliser la boîte de dialogue Office dans vos compléments Office](../develop/dialog-api-in-office-add-ins.md).

Ajoutez la fonction suivante à **./src/commands/commands.js**. Cette fonction définit la propriété personnalisée **EventLogged** sur l’élément de rendez-vous actuel.

```js
function updateCustomProperties() {
  Office.context.mailbox.item.loadCustomPropertiesAsync(
    function callback(customPropertiesResult) {
      if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
        let customProperties = customPropertiesResult.value;
        customProperties.set("EventLogged", true);
        customProperties.saveAsync(
          function callback(setSaveAsyncResult) {
            if (setSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("EventLogged custom property saved successfully.");
              event.completed({ allowEvent: true });
              event = undefined;
            }
          }
        );
      }
    }
  );
}
```

Appelez-le une fois que le complément a correctement journaliser les notes de rendez-vous. Par exemple, vous pouvez l’appeler à partir de **logCRMEvent** , comme indiqué dans la fonction suivante.

```js
function logCRMEvent(appointmentEvent) {
  event = appointmentEvent;
  console.log(`Subject: ${Office.context.mailbox.item.subject}`);
  Office.context.mailbox.item.body.getAsync(
    "html",
    { asyncContext: "This is passed to the callback" },
    function callback(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        // Replace `event.completed({ allowEvent: true });` with the following statement.
        updateCustomProperties();
      } else {
        console.error("Failed to get body.");
        event.completed({ allowEvent: false });
      }
    }
  );
}
```

### <a name="delete-the-appointment-log"></a>Supprimer le journal des rendez-vous

Si vous souhaitez permettre à vos utilisateurs d’annuler la journalisation ou de supprimer les notes de rendez-vous journalisées afin d’enregistrer un journal de remplacement, vous disposez de deux options.

1. Utilisez Microsoft Graph pour [effacer l’objet de propriétés personnalisées](/graph/api/resources/extended-properties-overview?view=graph-rest-1.0&preserve-view=true) lorsque l’utilisateur sélectionne le bouton approprié dans le ruban.
1. Ajoutez la fonction suivante à **./src/commands/commands.js** pour effacer la propriété personnalisée **EventLogged** sur l’élément de rendez-vous actuel.

    ```js
    function clearCustomProperties() {
      Office.context.mailbox.item.loadCustomPropertiesAsync(
        function callback(customPropertiesResult) {
          if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
            var customProperties = customPropertiesResult.value;
            customProperties.remove("EventLogged");
            customProperties.saveAsync(
              function callback(removeSaveAsyncResult) {
                if (removeSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log("Custom properties cleared");
                  event.completed({ allowEvent: true });
                  event = undefined;
                }
              }
            );
          }
        }
      );
    }
    ```

Appelez-la quand vous voulez effacer la propriété personnalisée. Par exemple, vous pouvez l’appeler à partir de **logCRMEvent** si la définition du journal a échoué d’une manière ou d’une autre, comme indiqué dans la fonction suivante.

  ```js
  function logCRMEvent(appointmentEvent) {
    event = appointmentEvent;
    console.log(`Subject: ${Office.context.mailbox.item.subject}`);
    Office.context.mailbox.item.body.getAsync(
      "html",
      { asyncContext: "This is passed to the callback" },
      function callback(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          updateCustomProperties();
        } else {
          console.error("Failed to get body.");
          // Replace `event.completed({ allowEvent: false });` with the following statement.
          clearCustomProperties();
        }
      }
    );
  }
  ```

# <a name="task-pane"></a>[Volet Office](#tab/taskpane)

Cette option permet à un utilisateur de journaliser et d’afficher ses notes et d’autres détails sur ses rendez-vous à partir d’un volet Office.

### <a name="configure-the-manifest"></a>Configurer le manifeste

Pour permettre aux utilisateurs de consigner les notes de rendez-vous avec votre complément, vous devez configurer le [point d’extension MobileLogEventAppointmentAttendee](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee) dans le manifeste sous l’élément `MobileFormFactor`parent. D’autres facteurs de forme ne sont pas pris en charge.

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

1. Dans votre éditeur de code, ouvrez le projet de démarrage rapide.

1. Ouvrez le fichier **manifest.xml** situé à la racine de votre projet.

1. Sélectionnez l’intégralité `<VersionOverrides>` du nœud (y compris les balises d’ouverture et de fermeture) et remplacez-le par le code XML suivant. Veillez à remplacer toutes les références à **Contoso** par les informations de votre entreprise.

    ```xml
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
      <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
        <Description resid="residDescription"></Description>
          <Requirements>
            <bt:Sets>
              <bt:Set Name="Mailbox" MinVersion="1.3"/>
            </bt:Sets>
          </Requirements>
          <Hosts>
            <Host xsi:type="MailHost">
              <DesktopFormFactor>
                <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
                  <OfficeTab id="TabDefault">
                    <Group id="apptReadGroup">
                      <Label resid="residDescription"/>
                      <Control xsi:type="Button" id="apptReadOpenPaneButton">
                        <Label resid="residLabel"/>
                        <Supertip>
                          <Title resid="residLabel"/>
                          <Description resid="residTooltip"/>
                        </Supertip>
                        <Icon>
                          <bt:Image size="16" resid="icon-16"/>
                          <bt:Image size="32" resid="icon-32"/>
                          <bt:Image size="80" resid="icon-80"/>
                        </Icon>
                        <Action xsi:type="ShowTaskpane">
                          <SourceLocation resid="Taskpane.Url"/>
                        </Action>
                      </Control>
                    </Group>
                  </OfficeTab>
                </ExtensionPoint>
              </DesktopFormFactor>
              <MobileFormFactor>
                <ExtensionPoint xsi:type="MobileLogEventAppointmentAttendee">
                  <Control xsi:type="MobileButton" id="appointmentReadFunctionButton">
                    <Label resid="residLabel"/>
                    <Icon>
                      <bt:Image size="25" scale="1" resid="icon-16"/>
                      <bt:Image size="25" scale="2" resid="icon-16"/>
                      <bt:Image size="25" scale="3" resid="icon-16"/>
    
                      <bt:Image size="32" scale="1" resid="icon-32"/>
                      <bt:Image size="32" scale="2" resid="icon-32"/>
                      <bt:Image size="32" scale="3" resid="icon-32"/>
    
                      <bt:Image size="48" scale="1" resid="icon-48"/>
                      <bt:Image size="48" scale="2" resid="icon-48"/>
                      <bt:Image size="48" scale="3" resid="icon-48"/>
                    </Icon>
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="Taskpane.Url"/>
                    </Action> 
                  </Control>
                </ExtensionPoint>
              </MobileFormFactor>
            </Host>
          </Hosts>
          <Resources>
            <bt:Images>
              <bt:Image id="icon-16" DefaultValue="https://contoso.com/assets/icon-16.png"/>
              <bt:Image id="icon-32" DefaultValue="https://contoso.com/assets/icon-32.png"/>
              <bt:Image id="icon-48" DefaultValue="https://contoso.com/assets/icon-48.png"/>
              <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
            </bt:Images>
            <bt:Urls>
              <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
              <bt:Url id="Taskpane.Url" DefaultValue="https://contoso.com/taskpane.html"/>
            </bt:Urls>
            <bt:ShortStrings>
              <bt:String id="residDescription" DefaultValue="Log appointment notes and other details to Contoso CRM."/>
              <bt:String id="residLabel" DefaultValue="Log to Contoso CRM"/>
            </bt:ShortStrings>
            <bt:LongStrings>
              <bt:String id="residTooltip" DefaultValue="Log notes to Contoso CRM for this appointment."/>
            </bt:LongStrings>
          </Resources>
        </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> Pour en savoir plus sur les manifestes pour les compléments Outlook, consultez [les manifestes de complément Outlook](manifests.md) et [la prise en charge des commandes de complément pour Outlook Mobile](add-mobile-support.md).

### <a name="capture-appointment-notes"></a>Capturer les notes de rendez-vous

Dans cette section, découvrez comment afficher les notes de rendez-vous journalisées et d’autres détails dans un volet Office lorsque l’utilisateur sélectionne le bouton **Journal** .

1. À partir du même projet de démarrage rapide, ouvrez le fichier **./src/taskpane/taskpane.js** dans votre éditeur de code.

1. Remplacez l’intégralité du contenu du fichier **taskpane.js** par le code JavaScript suivant.

    ```js
    // Office is ready.
    Office.onReady(function () {
        getEventData();
      }
    );

    function getEventData() {
      console.log(`Subject: ${Office.context.mailbox.item.subject}`);
      Office.context.mailbox.item.body.getAsync(
        "html",
        function callback(result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("event logged successfully");
          } else {
            console.error("Failed to get body.");
          }
        }
      );
    }
    ```

Ensuite, mettez à jour le fichier **taskpane.html** pour référencer **taskpane.js**.

1. À partir du même projet de démarrage rapide, ouvrez le fichier **./src/taskpane/taskpane.html** dans votre éditeur de code.

1. Recherchez et remplacez `<script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>` les éléments suivants :

    ```html
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
    <script type="text/javascript" src="taskpane.js"></script>
    ```

### <a name="view-appointment-notes"></a>Afficher les notes de rendez-vous

L’étiquette du bouton **Journal** peut être activée pour afficher **l’affichage** en définissant la propriété personnalisée **Journalisée des événements** réservée à cet effet. Lorsque l’utilisateur sélectionne le bouton **Afficher** , il peut consulter ses notes journalisées pour ce rendez-vous. Votre complément définit l’expérience d’affichage des journaux.

Ajoutez la fonction suivante à **./src/taskpane/taskpane.js**. Cette fonction définit la propriété personnalisée **EventLogged** sur l’élément de rendez-vous actuel.

```js
function updateCustomProperties() {
  Office.context.mailbox.item.loadCustomPropertiesAsync(
    function callback(customPropertiesResult) {
      if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
        let customProperties = customPropertiesResult.value;
        customProperties.set("EventLogged", true);
        customProperties.saveAsync(
          function callback(setSaveAsyncResult) {
            if (setSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("EventLogged custom property saved successfully.");
            }
          }
        );
      }
    }
  );
}
```

Appelez-le une fois que le complément a correctement journaliser les notes de rendez-vous. Par exemple, vous pouvez l’appeler à partir de **getEventData** , comme indiqué dans la fonction suivante.

```js
function getEventData() {
  console.log(`Subject: ${Office.context.mailbox.item.subject}`);
  Office.context.mailbox.item.body.getAsync(
    "html",
    function callback(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("event logged successfully");
        updateCustomProperties();
      } else {
        console.error("Failed to get body.");
      }
    }
  );
}
```

### <a name="delete-the-appointment-log"></a>Supprimer le journal des rendez-vous

Si vous souhaitez permettre à vos utilisateurs d’annuler la journalisation ou de supprimer les notes de rendez-vous journalisées afin d’enregistrer un journal de remplacement, vous disposez de deux options.

1. Utilisez Microsoft Graph pour [effacer l’objet de propriétés personnalisées](/graph/api/resources/extended-properties-overview?view=graph-rest-1.0&preserve-view=true) lorsque l’utilisateur sélectionne le bouton approprié dans le volet Office.
1. Ajoutez la fonction suivante à **./src/taskpane/taskpane.js** pour effacer la propriété personnalisée **EventLogged** sur l’élément de rendez-vous actuel.

    ```js
    function clearCustomProperties() {
      Office.context.mailbox.item.loadCustomPropertiesAsync(
        function callback(customPropertiesResult) {
          if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
            var customProperties = customPropertiesResult.value;
            customProperties.remove("EventLogged");
            customProperties.saveAsync(
              function callback(removeSaveAsyncResult) {
                if (removeSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log("Custom properties cleared");
                }
              }
            );
          }
        }
      );
    }
    ```

Appelez-la quand vous voulez effacer la propriété personnalisée. Par exemple, vous pouvez l’appeler à partir de **getEventData** si la définition du journal a échoué d’une manière ou d’une autre, comme indiqué dans la fonction suivante.

  ```js
  function getEventData() {
    console.log(`Subject: ${Office.context.mailbox.item.subject}`);
    Office.context.mailbox.item.body.getAsync(
      "html",
      function callback(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("event logged successfully");
          updateCustomProperties();
        } else {
          console.error("Failed to get body.");
          clearCustomProperties();
        }
      }
    );
  }
  ```

---

## <a name="test-and-validate"></a>Tester et valider

1. Suivez les instructions habituelles pour [tester et valider votre complément](testing-and-tips.md).
1. Après avoir [chargé de manière indépendante](sideload-outlook-add-ins-for-testing.md) le complément dans Outlook sur le web, Windows ou Mac, redémarrez Outlook sur votre appareil mobile Android.
1. Ouvrez un rendez-vous en tant que participant, puis vérifiez que sous la carte **Insights** de réunion, il existe une nouvelle carte avec le nom de votre complément à côté du bouton **Journal** .

### <a name="ui-log-the-appointment-notes"></a>Interface utilisateur : journaliser les notes de rendez-vous

En tant que participant à une réunion, vous devez voir un écran similaire à l’image suivante lorsque vous ouvrez une réunion.

![Capture d’écran montrant le bouton Journal sur un écran de rendez-vous sur Android.](../images/outlook-android-log-appointment-details.jpg)

### <a name="ui-view-the-appointment-log"></a>Interface utilisateur : afficher le journal des rendez-vous

Une fois les notes de rendez-vous correctement enregistrées, le bouton doit maintenant être étiqueté **Affichage** au lieu de **Journal**. Vous devez voir un écran similaire à l’image suivante.

![Capture d’écran montrant le bouton Afficher sur un écran de rendez-vous sur Android.](../images/outlook-android-view-appointment-log.jpg)

## <a name="available-apis"></a>API disponibles

Les API suivantes sont disponibles pour cette fonctionnalité.

- [API de boîte de dialogue](../develop/dialog-api-in-office-add-ins.md)
- [Office.AddinCommands.Event](/javascript/api/office/office.addincommands.event?view=outlook-js-preview&preserve-view=true)
- [Office.CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true)
- [Office.RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true)
- [API de lecture de rendez-vous (participant),](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true) **à l’exception** des suivantes :
  - [Office.context.mailbox.item.categories](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#categories)
  - [Office.context.mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#enhancedLocation)
  - [Office.context.mailbox.item.isAllDayEvent](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#isAllDayEvent)
  - [Office.context.mailbox.item.recurrence](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#recurrence)
  - [Office.context.mailbox.item.sensitivity](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#sensitivity)
  - [Office.context.mailbox.item.seriesId](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#seriesId)

## <a name="restrictions"></a>Restrictions

Plusieurs restrictions s’appliquent.

- Impossible **de** modifier le nom du bouton Journal. Toutefois, il existe un moyen d’afficher une autre étiquette en définissant une propriété personnalisée sur l’élément de rendez-vous. Pour plus d’informations, reportez-vous à la section **Afficher les notes de rendez-vous** pour la [commande de fonction](?tabs=noui#view-appointment-notes) ou [le volet Office](?tabs=taskpane#view-appointment-notes-1) , le cas échéant.
- La propriété personnalisée **EventLogged** doit être utilisée si vous souhaitez activer l’étiquette du bouton **Journal** pour **afficher** et revenir.
- L’icône de complément doit être en nuances de gris à l’aide de code `#919191` hexadécimal ou de son équivalent dans [d’autres formats de couleur](https://convertingcolors.com/hex-color-919191.html).
- Le complément doit extraire les détails de la réunion du formulaire de rendez-vous dans le délai d’expiration d’une minute. Toutefois, tout temps passé dans une boîte de dialogue le complément ouvert pour l’authentification, par exemple, est exclu du délai d’expiration.

## <a name="see-also"></a>Voir aussi

- [Compléments pour Outlook Mobile](outlook-mobile-addins.md)
- [Ajouter la prise en charge des commandes de complément pour Outlook Mobile](add-mobile-support.md)
