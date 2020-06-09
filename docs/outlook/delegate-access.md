---
title: Activer les scénarios d’accès délégué dans un complément Outlook
description: Décrit brièvement l’accès délégué et explique comment configurer la prise en charge des compléments.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 68b9e09afbe2bcd5cfc302d6714b1c22fd945047
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608949"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a>Activer les scénarios d’accès délégué dans un complément Outlook

Un propriétaire de boîte aux lettres peut utiliser la fonctionnalité accès délégué pour [permettre à quelqu’un d’autre de gérer son courrier et son calendrier](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926). Cet article indique les autorisations déléguées prises en charge par l’API JavaScript pour Office et explique comment activer les scénarios d’accès délégué dans votre complément Outlook.

> [!IMPORTANT]
> L’accès délégué n’est pas disponible actuellement dans Outlook sur Mac, Android et iOS. Cette fonctionnalité peut être rendue disponible à l’avenir.
>
> La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1,8. Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

## <a name="supported-permissions-for-delegate-access"></a>Autorisations prises en charge pour l’accès délégué

Le tableau suivant décrit les autorisations déléguées prises en charge par l’API JavaScript pour Office.

|Permission|Valeur|Description|
|---|---:|---|
|Read|1 (000001)|Peut lire des éléments.|
|Écriture|2 (000010)|Peut créer des éléments.|
|DeleteOwn|4 (000100)|Peut uniquement supprimer les éléments qu’ils ont créés.|
|DeleteAll|8 (001000)|Peut supprimer tous les éléments.|
|EditOwn|16 (010000)|Ne peut modifier que les éléments qu’ils ont créés.|
|EditAll|32 (100000)|Peut modifier tous les éléments.|

> [!NOTE]
> Actuellement, l’API prend en charge l’obtention des autorisations de délégué existantes, mais pas la définition des autorisations de délégué.

L’objet [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) est implémenté à l’aide d’un masque de masques pour indiquer les autorisations du délégué. Chaque position dans le masque de données représente une autorisation particulière et si elle est définie sur `1` Then, le délégué dispose de l’autorisation correspondante. Par exemple, si le deuxième bit à partir de la droite est `1` , le délégué dispose alors d’une autorisation en **écriture** . Vous pouvez voir un exemple sur la façon de vérifier une autorisation spécifique dans la section [effectuer une opération en tant que délégué](#perform-an-operation-as-delegate) plus loin dans cet article.

## <a name="sync-across-mailbox-clients"></a>Synchronisation entre les clients de boîte aux lettres

Les mises à jour d’un délégué vers la boîte aux lettres du propriétaire sont généralement synchronisées entre les boîtes aux lettres immédiatement.

Toutefois, si le complément utilise les opérations REST ou EWS pour définir une propriété étendue sur un élément, la synchronisation de ces modifications peut prendre quelques heures. Nous vous recommandons d’utiliser à la place l’objet [CustomProperties](/javascript/api/outlook/office.customproperties) et les API associées pour éviter ce délai. Pour plus d’informations, reportez-vous à la [section Propriétés personnalisées](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) de l’article « obtenir et définir des métadonnées dans un complément Outlook ».

## <a name="configure-the-manifest"></a>Configurer le manifeste

Pour activer les scénarios d’accès délégué dans votre complément, vous devez définir l’élément [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) `true` dans le manifeste sous l’élément parent `DesktopFormFactor` . Actuellement, les autres facteurs de forme ne sont pas pris en charge.

L’exemple suivant montre l' `SupportsSharedFolders` élément défini `true` dans une section du manifeste.

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

Vous pouvez obtenir les propriétés partagées d’un élément en mode de composition ou de lecture en appelant la méthode [Item. getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) . Cela renvoie un objet [SharedProperties](/javascript/api/outlook/office.sharedproperties) qui fournit actuellement les autorisations du délégué, l’adresse de messagerie du propriétaire, l’URL de base de l’API REST et la boîte aux lettres cible.

L’exemple suivant montre comment obtenir les propriétés partagées d’un message ou d’un rendez-vous, vérifier si le délégué dispose d’une autorisation en **écriture** et passer un appel Rest.

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

## <a name="see-also"></a>Voir aussi

- [Autoriser quelqu’un d’autre à gérer votre courrier et votre calendrier](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [Partage de calendriers dans Office 365](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Procédure de tri des éléments de manifeste](../develop/manifest-element-ordering.md)
- [Mask (Computing)](https://en.wikipedia.org/wiki/Mask_(computing))
- [Opérateurs de bits JavaScript](https://www.w3schools.com/js/js_bitwise.asp)