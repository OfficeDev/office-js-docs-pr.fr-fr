---
title: Activer les dossiers partagés et les scénarios de boîte aux lettres partagées dans un complément Outlook
description: Explique comment configurer la prise en charge des compléments pour les dossiers partagés (par exemple, déléguer l’accès) et aux boîtes aux lettres partagées.
ms.date: 09/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1c6884c18e4cb9916fcec20e6b732b0d20918e2f
ms.sourcegitcommit: 54a7dc07e5f31dd5111e4efee3e85b4643c4bef5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/21/2022
ms.locfileid: "67857556"
---
# <a name="enable-shared-folders-and-shared-mailbox-scenarios-in-an-outlook-add-in"></a>Activer les dossiers partagés et les scénarios de boîte aux lettres partagées dans un complément Outlook

Cet article explique comment activer des dossiers partagés (également appelés accès délégué) et des boîtes aux lettres partagées (désormais en [préversion](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview#shared-mailboxes)) dans votre complément Outlook, y compris les autorisations prises en charge par l’API JavaScript Office.

## <a name="supported-clients-and-platforms"></a>Plateformes et clients pris en charge

Le tableau suivant présente les combinaisons client-serveur prises en charge pour cette fonctionnalité, y compris la mise à jour cumulative minimale requise, le cas échéant. Les combinaisons exclues ne sont pas prises en charge.

| Client | Exchange Online | Exchange 2019 local<br>(Mise à jour cumulative 1 ou ultérieure) | Exchange 2016 local<br>(Mise à jour cumulative 6 ou ultérieure) | Exchange 2013 local |
|---|:---:|:---:|:---:|:---:|
|Windows :<br>version 1910 (build 12130.20272) ou ultérieure|Oui|Oui\*|Oui\*|Oui\*|
|Mac:<br>build 16.47 ou ultérieure|Oui|Oui|Oui|Oui|
|Navigateur web :<br>interface utilisateur Outlook moderne|Oui|Non applicable|Non applicable|Non applicable|
|Navigateur web :<br>Interface utilisateur Outlook classique|Non applicable|Non|Non|Non|

> [!NOTE]
> \* La prise en charge de cette fonctionnalité dans un environnement Exchange local est disponible à partir de la version 2206 (build 15330.20000) pour le canal actuel et de la version 2207 (build 15427.20000) pour le canal Entreprise mensuel.

> [!IMPORTANT]
> La prise en charge de cette fonctionnalité a été introduite dans [l’ensemble de conditions requises 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) (pour plus d’informations, reportez-vous aux [clients et aux plateformes](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)). Toutefois, notez que la matrice de prise en charge de la fonctionnalité est un sur-ensemble de l’ensemble de conditions requises.

## <a name="supported-setups"></a>Configurations prises en charge

Les sections suivantes décrivent les configurations prises en charge pour les boîtes aux lettres partagées (maintenant en préversion) et les dossiers partagés. Les API de fonctionnalité peuvent ne pas fonctionner comme prévu dans d’autres configurations. Sélectionnez la plateforme que vous souhaitez savoir comment configurer.

### <a name="windows"></a>[Fenêtres](#tab/windows)

#### <a name="shared-folders"></a>Dossiers partagés

Le propriétaire de la boîte aux lettres doit [d’abord fournir l’accès à un délégué](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926). Le délégué doit ensuite suivre les instructions décrites dans la section « Ajouter la boîte aux lettres d’une autre personne à votre profil » de [l’article Gérer les éléments de courrier et de calendrier d’une autre personne](https://support.microsoft.com/office/afb79d6b-2967-43b9-a944-a6b953190af5).

#### <a name="shared-mailboxes-preview"></a>Boîtes aux lettres partagées (préversion)

Les administrateurs de serveur Exchange peuvent créer et gérer des boîtes aux lettres partagées aux fins d’accès à des ensembles d’utilisateurs. [les environnements Exchange Exchange Online](/exchange/collaboration-exo/shared-mailboxes) et [locaux](/exchange/collaboration/shared-mailboxes/create-shared-mailboxes) sont pris en charge.

Une fonctionnalité de Exchange Server appelée « mise en page automatique » est activée par défaut, ce qui signifie que par la suite, la [boîte aux lettres partagée doit apparaître automatiquement](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) dans l’application Outlook d’un utilisateur après la fermeture et la réouverture d’Outlook. Toutefois, si un administrateur a désactivé le mappage automatique, l’utilisateur doit suivre les étapes manuelles décrites dans la section « Ajouter une boîte aux lettres partagée à Outlook » de l’article [Ouvrir et utiliser une boîte aux lettres partagée dans Outlook](https://support.microsoft.com/office/d94a8e9e-21f1-4240-808b-de9c9c088afd).

> [!WARNING]
> Ne **vous** connectez PAS à la boîte aux lettres partagée avec un mot de passe. Les API de fonctionnalité ne fonctionneront pas dans ce cas.

### <a name="web-browser---modern-outlook"></a>[Navigateur web – Outlook moderne](#tab/modern)

#### <a name="shared-folders"></a>Dossiers partagés

Le propriétaire de la boîte aux lettres doit [d’abord fournir l’accès à un délégué](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) en mettant à jour les autorisations du dossier de boîte aux lettres. Le délégué doit ensuite suivre les instructions décrites dans la section « Ajouter la boîte aux lettres d’une autre personne à votre liste de dossiers dans Outlook Web App » de l’article [Accéder à la boîte aux lettres d’une autre personne](https://support.microsoft.com/office/a909ad30-e413-40b5-a487-0ea70b763081).

#### <a name="shared-mailboxes"></a>Boîtes aux lettres partagées

Les scénarios de boîte aux lettres partagées dans les compléments Outlook ne sont actuellement pas pris en charge dans les Outlook sur le web modernes.

### <a name="mac"></a>[Mac](#tab/unix)

#### <a name="shared-mailboxes-preview"></a>Boîtes aux lettres partagées (préversion)

Le courrier et le calendrier sont partagés avec un délégué ou un utilisateur de boîte aux lettres partagé. Les compléments sont disponibles pour le délégué ou l’utilisateur dans les modes de lecture et de composition des messages et des rendez-vous.

#### <a name="shared-folders"></a>Dossiers partagés

Si le dossier **boîte de réception** est partagé avec un délégué, les compléments sont disponibles pour le délégué en mode de lecture des messages.

Si le dossier **Brouillons** est également partagé avec le délégué, les compléments sont disponibles en mode composition.

#### <a name="local-shared-calendar-new-model"></a>Calendrier partagé local (nouveau modèle)

Si le propriétaire du calendrier a explicitement partagé son calendrier avec un délégué (la boîte aux lettres entière ne peut pas être partagée), les compléments sont disponibles pour le délégué en mode lecture et composition du rendez-vous.

#### <a name="remote-shared-calendar-previous-model"></a>Calendrier partagé distant (modèle précédent)

Si le propriétaire du calendrier a accordé un accès étendu à son calendrier (par exemple, l’a rendu modifiable pour une DL particulière ou l’ensemble de l’organisation), les utilisateurs peuvent alors disposer d’autorisations indirectes ou implicites et les compléments sont disponibles pour ces utilisateurs en mode lecture et composition de rendez-vous.

---

Pour en savoir plus sur l’endroit où les compléments effectuent et ne s’activent pas en général, reportez-vous à la section Éléments de boîte aux [lettres disponibles pour les compléments](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) de la page vue d’ensemble des compléments Outlook.

## <a name="supported-permissions"></a>Autorisations prises en charge

Le tableau suivant décrit les autorisations prises en charge par l’API JavaScript Office pour les délégués et les utilisateurs de boîtes aux lettres partagées.

|Autorisation|Valeur|Description|
|---|---:|---|
|Lecture|1 (000001)|Peut lire des éléments.|
|Write|2 (000010)|Peut créer des éléments.|
|DeleteOwn|4 (000100)|Peut supprimer uniquement les éléments qu’ils ont créés.|
|DeleteAll|8 (001000)|Peut supprimer tous les éléments.|
|EditOwn|16 (010000)|Peut modifier uniquement les éléments qu’ils ont créés.|
|EditAll|32 (100000)|Peut modifier tous les éléments.|

> [!NOTE]
> Actuellement, l’API prend en charge l’obtention d’autorisations existantes, mais pas la définition d’autorisations.

L’objet [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) est implémenté à l’aide d’un masque de bits pour indiquer les autorisations. Chaque position dans le masque de bits représente une autorisation particulière et, si elle est définie sur `1` , l’utilisateur dispose de l’autorisation correspondante. Par exemple, si le deuxième bit à partir de la droite est `1`, l’utilisateur dispose de l’autorisation **d’écriture** . Vous pouvez voir un exemple montrant comment rechercher une autorisation spécifique dans la section [Effectuer une opération en tant que délégué ou utilisateur de boîte aux lettres partagée](#perform-an-operation-as-delegate-or-shared-mailbox-user) plus loin dans cet article.

## <a name="sync-across-shared-folder-clients"></a>Synchroniser entre les clients de dossiers partagés

Les mises à jour d’un délégué dans la boîte aux lettres du propriétaire sont généralement synchronisées immédiatement entre les boîtes aux lettres.

Toutefois, si les opérations REST ou Exchange Web Services (EWS) ont été utilisées pour définir une propriété étendue sur un élément, la synchronisation de ces modifications peut prendre quelques heures. Nous vous recommandons plutôt d’utiliser l’objet [CustomProperties](/javascript/api/outlook/office.customproperties) et les API associées pour éviter un tel délai. Pour plus d’informations, consultez la [section propriétés personnalisées](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) de l’article « Obtenir et définir des métadonnées dans un complément Outlook ».

> [!IMPORTANT]
> Dans un scénario de délégué, vous ne pouvez pas utiliser EWS avec les jetons actuellement fournis par office.js API.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Pour activer les dossiers partagés et les scénarios de boîte aux lettres partagées dans votre complément, vous devez définir l’élément `true` [SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) sur dans le manifeste sous l’élément `DesktopFormFactor`parent. À l’heure actuelle, d’autres facteurs de forme ne sont pas pris en charge.

Pour prendre en charge les appels REST à partir d’un délégué, définissez le nœud [Autorisations](/javascript/api/manifest/permissions) dans le manifeste `ReadWriteMailbox`sur .

L’exemple suivant montre l’élément `SupportsSharedFolders` défini `true` dans une section du manifeste.

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

## <a name="perform-an-operation-as-delegate-or-shared-mailbox-user"></a>Effectuer une opération en tant qu’utilisateur délégué ou de boîte aux lettres partagée

Vous pouvez obtenir les propriétés partagées d’un élément en mode Compose ou Lecture en appelant la méthode [item.getSharedPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) . Cela renvoie un objet [SharedProperties](/javascript/api/outlook/office.sharedproperties) qui fournit actuellement les autorisations de l’utilisateur, l’adresse e-mail du propriétaire, l’URL de base de l’API REST et la boîte aux lettres cible.

L’exemple suivant montre comment obtenir les propriétés partagées d’un message ou d’un rendez-vous, vérifier si le délégué ou l’utilisateur de boîte aux lettres partagée dispose d’une autorisation **d’écriture** et effectuer un appel REST.

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
> En tant que délégué, vous pouvez utiliser REST pour [obtenir le contenu d’un message Outlook attaché à un élément Outlook ou à un billet de groupe](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a>Gérer l’appel de REST sur des éléments partagés et non partagés

Si vous souhaitez appeler une opération REST sur un élément, que l’élément soit partagé ou non, vous pouvez utiliser l’API `getSharedPropertiesAsync` pour déterminer si l’élément est partagé. Après cela, vous pouvez construire l’URL REST pour l’opération à l’aide de l’objet approprié.

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://learn.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## <a name="limitations"></a>Limites

Selon les scénarios de votre complément, il existe quelques limitations à prendre en compte lors de la gestion de dossiers partagés ou de boîtes aux lettres partagées.

### <a name="message-compose-mode"></a>Mode Composition des messages

En mode Composition de message, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-getsharedpropertiesasync-member(1)) n’est pas pris en charge dans Outlook sur le web ou sur Windows, sauf si les conditions suivantes sont remplies.

a. **Déléguer l’accès/dossiers partagés**

1. Le propriétaire de la boîte aux lettres démarre un message. Il peut s’agir d’un nouveau message, d’une réponse ou d’un transfert.
1. Ils enregistrent le message, puis le déplacent de leur propre dossier **Brouillons** vers un dossier partagé avec le délégué.
1. Le délégué ouvre le brouillon à partir du dossier partagé, puis continue à composer.

b. **Boîte aux lettres partagée (s’applique à Outlook sur Windows uniquement)**

1. Un utilisateur de boîte aux lettres partagée démarre un message. Il peut s’agir d’un nouveau message, d’une réponse ou d’un transfert.
1. Ils enregistrent le message, puis le déplacent de leur propre dossier **Brouillons** vers un dossier dans la boîte aux lettres partagée.
1. Un autre utilisateur de boîte aux lettres partagée ouvre le brouillon à partir de la boîte aux lettres partagée, puis continue à composer.

Le message se trouve maintenant dans un contexte partagé et les compléments qui prennent en charge ces scénarios partagés peuvent obtenir les propriétés partagées de l’élément. Une fois le message envoyé, il se trouve généralement dans le dossier **Éléments envoyés** de l’expéditeur.

### <a name="rest-and-ews"></a>REST et EWS

Votre complément peut utiliser REST et l’autorisation du complément doit être définie pour `ReadWriteMailbox` activer l’accès REST à la boîte aux lettres du propriétaire ou à la boîte aux lettres partagée, le cas échéant. EWS n’est pas pris en charge.

### <a name="user-or-shared-mailbox-hidden-from-an-address-list"></a>Boîte aux lettres utilisateur ou partagée masquée dans une liste d’adresses

Si un administrateur masquait une adresse de boîte aux lettres utilisateur ou partagée à partir d’une liste d’adresses telle que la liste d’adresses globale (GAL), les éléments de messagerie affectés ouverts dans le rapport `Office.context.mailbox.item` de boîte aux lettres sont null. Par exemple, si l’utilisateur ouvre un élément de messagerie dans une boîte aux lettres partagée qui est masquée de la gal, `Office.context.mailbox.item` ce qui représente cet élément de courrier est null.

## <a name="see-also"></a>Voir aussi

- [Autoriser une autre personne à gérer votre courrier électronique et votre calendrier](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926)
- [Partage de calendrier dans Microsoft 365](https://support.microsoft.com/office/b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Ajouter une boîte aux lettres partagée à Outlook](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [Guide pratique pour classer les éléments de manifeste](../develop/manifest-element-ordering.md)
- [Masque (informatique)](https://en.wikipedia.org/wiki/Mask_(computing))
- [Opérateurs de bits JavaScript](https://www.w3schools.com/js/js_bitwise.asp)
