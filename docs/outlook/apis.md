---
title: API de complément Outlook
description: Découvrez comment faire référence aux API de complément Outlook et déclarer des autorisations dans votre complément Outlook.
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 69043646add5e32502efb0d2a5b1259667e564d9
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467075"
---
# <a name="outlook-add-in-apis"></a>API de complément Outlook

Pour utiliser des API dans votre complément Outlook, vous devez spécifier l’emplacement de la bibliothèque Office.js, l’ensemble des conditions requises, le schéma et les autorisations. Vous allez principalement utiliser les API JavaScript Office exposées via l’objet [Mailbox](#mailbox-object) .

## <a name="officejs-library"></a>Bibliothèque Office.js

Pour interagir avec [l’API de complément Outlook](/javascript/api/outlook), vous devez utiliser les API JavaScript dans Office.js. Le réseau de distribution de contenu (CDN) de la bibliothèque est `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`. Les compléments soumis à AppSource doivent faire référence à Office.js par le biais de ce CDN et ne peuvent pas utiliser de référence locale.

Référencez le CDN dans une `<script>`balise`<head>` de la page web (fichier .html, .aspx ou .php) qui implémente l’interface utilisateur de votre complément.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

As we add new APIs, the URL to Office.js will stay the same. We will change the version in the URL only if we break an existing API behavior.

> [!IMPORTANT]
> Lors du développement d’un complément pour une application cliente Office, référencez l’API JavaScript Office à partir de la `<head>` section de la page. Ainsi, l’API est entièrement initialisée avant les éléments Body.

## <a name="requirement-sets"></a>Ensembles de conditions requises

Toutes les API Outlook appartiennent à [l’ensemble de conditions requises de boîte aux lettres](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets). L’ensemble de conditions requises `Mailbox` possède plusieurs versions, et chaque nouvel ensemble d’API publié appartient à une version supérieure de l’ensemble. Tous les clients Outlook ne prendront pas en charge l’ensemble d’API le plus récent lors de sa publication, mais si un client Outlook prend en charge un ensemble de conditions requises, toutes les API comprises dans cet ensemble seront également prises en charge.

To control which Outlook clients the add-in appears in, specify a minimum requirement set version in the manifest. For example, if you specify requirement set version 1.3, the add-in will not show up in any Outlook client that doesn't support a minimum version of 1.3.

Specifying a requirement set doesn't limit your add-in to the APIs in that version. If the add-in specifies requirement set v1.1 but is running in an Outlook client that supports v1.3, the add-in can still use v1.3 APIs. The requirement set only controls which Outlook clients the add-in appears in.

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

Votre complément requiert les autorisations appropriées pour utiliser les API dont il a besoin. En général, vous devez spécifier l’autorisation minimum nécessaire pour votre complément.

Il existe quatre niveaux d’autorisations ; **restreint**, **élément de lecture**, **élément en lecture/écriture** et **boîte aux lettres en lecture/écriture**. Pour plus d’informations. Pour plus de détails, voir [Présentation des autorisations du complément Outlook](understanding-outlook-add-in-permissions.md).

## <a name="mailbox-object"></a>Objet Mailbox

[!include[information about Mailbox object](../includes/mailbox-object-desc.md)]

## <a name="see-also"></a>Voir aussi

- [Manifestes de complément Outlook](manifests.md)
- [Présentation de l’ensemble de conditions requises pour les API Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)
- [Présentation des autorisations de complément Outlook](understanding-outlook-add-in-permissions.md).
- [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md)
