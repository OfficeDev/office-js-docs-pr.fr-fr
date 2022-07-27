---
title: API de complément Outlook
description: Découvrez comment faire référence aux API de complément Outlook et déclarer des autorisations dans votre complément Outlook.
ms.date: 07/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: bd0bcdd3dfac6def9443b09d9797bfd0667c3b3d
ms.sourcegitcommit: 15714ef1118083032e640413ede69a162c43daed
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/27/2022
ms.locfileid: "67037814"
---
# <a name="outlook-add-in-apis"></a>API de complément Outlook

Pour utiliser des API dans votre complément Outlook, vous devez spécifier l’emplacement de la bibliothèque Office.js, l’ensemble des conditions requises, le schéma et les autorisations. Vous allez principalement utiliser les API JavaScript Office exposées via l’objet [Mailbox](#mailbox-object) .

## <a name="officejs-library"></a>Bibliothèque Office.js

Pour interagir avec [l’API de complément Outlook](/javascript/api/outlook), vous devez utiliser les API JavaScript dans Office.js. Le réseau de distribution de contenu (CDN) de la bibliothèque est `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`. Les compléments soumis à AppSource doivent faire référence à Office.js par le biais de ce CDN et ne peuvent pas utiliser de référence locale.

Référencez le CDN dans une `<script>`balise`<head>` de la page web (fichier .html, .aspx ou .php) qui implémente l’interface utilisateur de votre complément.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

L’ajout de nouvelles API ne modifie pas l’URL vers Office.js. La version de l’URL sera modifiée uniquement si un comportement d’API existant est interrompu.

> [!IMPORTANT]
> Lors du développement d’un complément pour une application cliente Office, référencez l’API JavaScript Office à partir de la `<head>` section de la page. Ainsi, l’API est entièrement initialisée avant les éléments Body.

## <a name="requirement-sets"></a>Ensembles de conditions requises

Toutes les API Outlook appartiennent à [l’ensemble de conditions requises de boîte aux lettres](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets). L’ensemble de conditions requises `Mailbox` possède plusieurs versions, et chaque nouvel ensemble d’API publié appartient à une version supérieure de l’ensemble. Tous les clients Outlook ne prendront pas en charge l’ensemble d’API le plus récent lors de sa publication, mais si un client Outlook prend en charge un ensemble de conditions requises, toutes les API comprises dans cet ensemble seront également prises en charge.

Pour savoir dans quels clients Outlook le complément s’affiche, indiquez la version de l’ensemble de conditions requises dans le manifeste. Par exemple, si vous indiquez la version 1.3 de l’ensemble de conditions requises, le complément n’apparaîtra pas dans les clients Outlook qui ne prennent pas en charge la version minimale 1.3.

Le fait de spécifier un ensemble de conditions requises ne limite pas votre complément aux API de cette version. Si le complément spécifie l’ensemble de conditions requises version 1.1, mais est exécuté dans un client Outlook prenant en charge la version 1.3, le complément peut toujours utiliser les API v1.3. L’ensemble de conditions requises contrôle uniquement les clients Outlook dans lesquels le complément apparaît.

Pour vérifier la disponibilité des API à partir d’un ensemble de conditions requises de version supérieure à celle spécifiée dans le manifeste, vous pouvez utiliser l’API JavaScript standard :

```js
if (item.somePropertyOrFunction) {
   item.somePropertyOrFunction...  
}
```

> [!NOTE]
> Ces vérifications ne sont pas nécessaires pour les API appartenant à l’ensemble de conditions requises dont la version est la même que celle spécifiée dans le manifeste.

Spécifiez l’ensemble de conditions requises minimal prenant en charge l’ensemble d’API critique pour votre scénario, sans lequel les fonctionnalités de votre complément ne fonctionneront pas. Vous spécifiez l’ensemble de conditions requises dans le manifeste. Le balisage varie en fonction du manifeste que vous utilisez. 

- **Manifeste XML** : utilisez l’élément **\<Requirements\>** . Notez que l’élément **\<Methods\>** enfant de **\<Requirements\>** n’est pas pris en charge dans les compléments Outlook. Vous ne pouvez donc pas déclarer la prise en charge de méthodes spécifiques.
- **Manifeste Teams (préversion)** : utilisez la propriété « extensions.capabilities ». 

Pour plus d’informations, consultez [les manifestes de complément Outlook](manifests.md) et comprendre les [ensembles de conditions requises de l’API Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets).

## <a name="permissions"></a>Autorisations

Votre complément requiert les autorisations appropriées pour utiliser les API dont il a besoin. En général, vous devez spécifier l’autorisation minimum nécessaire pour votre complément. Les autorisations sont déclarées dans le manifeste. Le balisage varie en fonction du type de manifeste.

- **Manifeste XML** : utilisez l’élément **\<Permissions\>** .
- **Manifeste Teams (préversion)** : utilisez la propriété « authorization.permissions.resourceSpecific ». 

Il existe quatre niveaux d’autorisations possibles. Pour plus de détails, voir [Présentation des autorisations du complément Outlook](understanding-outlook-add-in-permissions.md).

<br/>

|Niveau d’autorisation</br>Nom du manifeste XML|Niveau d’autorisation</br>Nom du manifeste Teams|Description|
|:-----|:-----|:-----|
| **Restricted** | **MailboxItem.Restricted.User** | Permet l’utilisation d’entités, mais pas d’expressions régulières. |
| **ReadItem** | **MailboxItem.Read.User** | En plus des autorisations indiquées dans **Restricted**, il autorise :<ul><li>expressions régulières</li><li>l’accès en lecture de l’API du complément Outlook</li><li>l’obtention des propriétés de l’élément et du jeton de rappel</li></ul> |
| **ReadWriteItem** | **MailboxItem.ReadWrite.User** | En plus de ce qui est autorisé dans **ReadItem**, il autorise :<ul><li>l’accès total à l’API du complément Outlook, à l’exception de `makeEwsRequestAsync`</li><li>la définition des propriétés de l’élément</li></ul> |
| **ReadWriteMailbox** | **Mailbox.ReadWrite.User** | En plus de ce qui est autorisé dans **ReadWriteItem**, il autorise :<ul><li>la création, la lecture, l’écriture d’éléments et de dossiers</li><li>l’envoi d’éléments</li><li>l’appel de [makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)</li></ul> |

> [!NOTE]
> Une autorisation supplémentaire est nécessaire pour les compléments qui utilisent la fonctionnalité d’ajout sur envoi. Avec le manifeste XML, vous spécifiez l’autorisation dans l’élément [ExtendedPermissions](/javascript/api/manifest/extendedpermissions) . Pour plus d’informations, consultez [Implémenter l’ajout à l’envoi dans votre complément Outlook](append-on-send.md). Avec le manifeste Teams (préversion), vous spécifiez cette autorisation avec le nom **Mailbox.AppendOnSend.User** dans un objet supplémentaire dans le tableau « authorization.permissions.resourceSpecific ».

Pour plus d’informations, consultez la rubrique [Manifestes des compléments Outlook](manifests.md). Pour plus d’informations sur les problèmes de sécurité, consultez [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md).

## <a name="mailbox-object"></a>Objet Mailbox

[!include[information about Mailbox object](../includes/mailbox-object-desc.md)]

## <a name="see-also"></a>Voir aussi

- [Manifestes de complément Outlook](manifests.md)
- [Présentation de l’ensemble de conditions requises pour les API Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)
- [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md)
