---
title: Activer les scénarios d’accès délégué dans un add-in Outlook
description: Décrit brièvement l’accès délégué et explique comment configurer la prise en charge des add-ins.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 598f931dbf3a4be8adf029838084ec0767bf6518
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234239"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a>Activer les scénarios d’accès délégué dans un add-in Outlook

Un propriétaire de boîte aux lettres peut utiliser la fonctionnalité d’accès délégué pour permettre à quelqu’un d’autre de [gérer son courrier et son calendrier.](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926) Cet article spécifie les autorisations déléguées que l’API JavaScript Office prend en charge et décrit comment activer les scénarios d’accès délégué dans votre application Outlook.

> [!IMPORTANT]
> L’accès délégué n’est pas disponible actuellement dans Outlook sur Android et iOS. En outre, cette fonctionnalité n’est pas disponible actuellement avec les boîtes aux lettres [partagées de](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) groupe dans Outlook sur le web. Cette fonctionnalité peut être rendue disponible à l’avenir.
>
> La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1.8. Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

## <a name="supported-permissions-for-delegate-access"></a>Autorisations prise en charge pour l’accès délégué

Le tableau suivant décrit les autorisations déléguées que l’API JavaScript Office prend en charge.

|Autorisation|Valeur|Description|
|---|---:|---|
|Lire|1 (000001)|Peut lire des éléments.|
|Write|2 (000010)|Peut créer des éléments.|
|DeleteOwn|4 (000100)|Peut supprimer uniquement les éléments qu’ils ont créés.|
|DeleteAll|8 (001000)|Peut supprimer tous les éléments.|
|EditOwn|16 (010000)|Peut modifier uniquement les éléments qu’ils ont créés.|
|EditAll|32 (100000)|Peut modifier n’importe quel objet.|

> [!NOTE]
> Actuellement, l’API prend en charge l’obtention d’autorisations déléguées existantes, mais pas la définition d’autorisations déléguées.

[L’objet DelegatePermissions est](/javascript/api/outlook/office.mailboxenums.delegatepermissions) implémenté à l’aide d’un masque de bits pour indiquer les autorisations du délégué. Chaque position dans le masque de bits représente une autorisation particulière et, si elle est définie sur, le délégué dispose de `1` l’autorisation respective. Par exemple, si le deuxième bit à partir de la droite est `1` , le délégué dispose de **l’autorisation d’écriture.** Vous pouvez voir un exemple de vérification d’une autorisation spécifique dans la [section](#perform-an-operation-as-delegate) Effectuer une opération en tant que délégué plus loin dans cet article.

## <a name="sync-across-mailbox-clients"></a>Synchronisation entre les clients de boîte aux lettres

Les mises à jour d’un délégué vers la boîte aux lettres du propriétaire sont généralement synchronisées immédiatement entre les boîtes aux lettres.

Toutefois, si des opérations REST ou services web Exchange (EWS) ont été utilisées pour définir une propriété étendue sur un élément, la synchronisation de ces modifications peut prendre quelques heures. Nous vous recommandons plutôt d’utiliser [l’objet CustomProperties](/javascript/api/outlook/office.customproperties) et les API associées pour éviter ce délai. Pour en savoir plus, consultez la [section des](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) propriétés personnalisées de l’article « Obtenir et définir des métadonnées dans un add-in Outlook ».

> [!IMPORTANT]
> Dans un scénario de délégué, vous ne pouvez pas utiliser EWS avec les jetons actuellement fournis par office.js API.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Pour activer les scénarios d’accès délégué dans votre add-in, vous devez définir l’élément [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) dans le manifeste sous `true` l’élément `DesktopFormFactor` parent. Pour l’instant, les autres facteurs de forme ne sont pas pris en charge.

Pour prendre en charge les appels REST d’un délégué, définissez le nœud [Autorisations](../reference/manifest/permissions.md) dans le manifeste sur `ReadWriteMailbox` .

L’exemple suivant montre `SupportsSharedFolders` l’ensemble `true` d’éléments dans une section du manifeste.

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="perform-an-operation-as-delegate"></a>Effectuer une opération en tant que délégué

Vous pouvez obtenir les propriétés partagées d’un élément en mode Composition ou Lecture en appelant la méthode [item.getSharedPropertiesAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) Cela renvoie un [objet SharedProperties](/javascript/api/outlook/office.sharedproperties) qui fournit actuellement les autorisations du délégué, l’adresse e-mail du propriétaire, l’URL de base de l’API REST et la boîte aux lettres cible.

L’exemple suivant montre comment obtenir les propriétés partagées d’un message ou d’un rendez-vous, vérifier si le délégué dispose d’une autorisation d’écriture et passer un appel REST. 

```js
function performOperation() {
  Office.context.mailbox.getCallbackTokenAsync({
      isRest: true
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value !== "") {
        Office.context.mailbox.item.getSharedPropertiesAsync({
            // Pass auth token along.
            asyncContext: asyncResult.value
          },
          function (asyncResult1) {
            let sharedProperties = asyncResult1.value;
            let delegatePermissions = sharedProperties.delegatePermissions;

            // Determine if user can do the expected operation.
            // E.g., do they have Write permission?
            if ((delegatePermissions & Office.MailboxEnums.DelegatePermissions.Write) != 0) {
              // Construct REST URL for your operation.
              // Update <version> placeholder with actual Outlook REST API version e.g. "v2.0".
              // Update <operation> placeholder with actual operation.
              let rest_url = sharedProperties.targetRestUrl + "/<version>/users/" + sharedProperties.targetMailbox + "/<operation>";
  
              $.ajax({
                  url: rest_url,
                  dataType: 'json',
                  headers:
                  {
                    "Authorization": "Bearer " + asyncResult1.asyncContext
                  }
                }
              ).done(
                function (response) {
                  console.log("success");
                }
              ).fail(
                function (error) {
                  console.log("error message");
                }
              );
            }
          }
        );
      }
    }
  );
}
```

> [!TIP]
> En tant que délégué, vous pouvez utiliser REST pour obtenir le contenu d’un message Outlook joint à un élément Outlook ou à un [billet de groupe.](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a>Gérer l’appel de REST sur les éléments partagés et non partagés

Si vous souhaitez appeler une opération REST sur un élément, que l’élément soit partagé ou non, vous pouvez utiliser l’API pour déterminer si l’élément `getSharedPropertiesAsync` est partagé. Après cela, vous pouvez construire l’URL REST pour l’opération à l’aide de l’objet approprié.

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://docs.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## <a name="limitations"></a>Limites

Selon les scénarios de votre add-in, vous devez tenir compte de quelques limitations lors de la gestion des situations de délégué.

### <a name="rest-and-ews"></a>REST et EWS

Votre add-in peut utiliser REST, mais pas EWS, et l’autorisation du module doit être définie pour activer l’accès REST à la boîte aux lettres `ReadWriteMailbox` du propriétaire.

### <a name="message-compose-mode"></a>Mode composition de message

En mode composition de message, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getsharedpropertiesasync-options--callback-) n’est pas pris en charge dans Outlook sur le web ou Windows, sauf si les conditions suivantes sont remplies.

1. Le propriétaire partage au moins un dossier de boîte aux lettres avec le délégué.
1. Le délégué rédige un message dans le dossier partagé.

    Exemples :

    - Le délégué répond ou envoie un message électronique dans le dossier partagé.
    - Le délégué enregistre un brouillon, puis le déplace de son propre dossier **Brouillons** vers le dossier partagé. Le délégué ouvre le brouillon à partir du dossier partagé, puis continue la composition.

Une fois le message envoyé, il se trouve généralement dans le dossier Éléments envoyés **du** délégué.

## <a name="see-also"></a>Voir aussi

- [Autoriser quelqu’un d’autre à gérer votre courrier et votre calendrier](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [Partage de calendrier dans Microsoft 365](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Comment commander des éléments de manifeste](../develop/manifest-element-ordering.md)
- [Masque (calcul)](https://en.wikipedia.org/wiki/Mask_(computing))
- [Opérateurs de bits JavaScript](https://www.w3schools.com/js/js_bitwise.asp)