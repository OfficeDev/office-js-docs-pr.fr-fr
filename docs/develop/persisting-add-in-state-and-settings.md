---
title: Conserver l’état et les paramètres du complément
description: Apprenez à conserver les données dans les applications web de complément Office exécutées dans l’environnement sans état d’un contrôle de navigateur.
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: e2018e5ecf419744257cdceac31b8b1688fa65ff
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810007"
---
# <a name="persist-add-in-state-and-settings"></a>Conserver l’état et les paramètres du complément

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location.
To do that, you can:

- Utilisez les membres de l’API JavaScript Office qui stockent des données comme suit :
  - Paires nom/valeur dans un conteneur de propriétés stocké dans un emplacement qui dépend du type de complément.
  - Éléments XML personnalisés stockés dans le document.

- Utilisez des techniques fournies par le contrôle de navigateur sous-jacent : les cookies de navigateur ou le stockage web HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) ou [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).
    > [!NOTE]
    > Certains navigateurs ou les paramètres du navigateur de l’utilisateur peuvent bloquer les techniques de stockage basées sur le navigateur. Vous devez tester la disponibilité comme indiqué dans [Utilisation de l’API de stockage web](https://developer.mozilla.org/docs/Web/API/Web_Storage_API/Using_the_Web_Storage_API).

Cet article se concentre sur l’utilisation de l’API JavaScript Office pour conserver l’état du complément dans le document actuel. Si vous avez besoin de conserver l’état sur les documents, comme le suivi des préférences utilisateur dans tous les documents qu’ils ouvrent, vous devez utiliser une approche différente. Par exemple, vous pouvez utiliser [l’authentification unique](use-sso-to-get-office-signed-in-user-token.md) pour obtenir l’identité de l’utilisateur, puis enregistrer l’ID utilisateur et ses paramètres dans une base de données en ligne.

## <a name="persist-add-in-state-and-settings-with-the-office-javascript-api"></a>Conserver l’état et les paramètres du complément avec l’API JavaScript Office

L’API JavaScript Office fournit les objets [Settings](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings) et [CustomProperties](/javascript/api/outlook/office.customproperties) pour enregistrer l’état du complément entre les sessions, comme décrit dans le tableau suivant. Dans tous les cas, les valeurs de paramètre enregistrées sont associées à l’[ID](/javascript/api/manifest/id) du complément qui les a créées.

|Objet|Prise en charge du type de complément|Emplacement de stockage|Prise en charge des applications Office|
|:-----|:-----|:-----|:-----|
|[Paramètres](/javascript/api/office/office.settings)|-Contenu<br>- Volet Office|Document, feuille de calcul ou présentation qu’utilise le complément. Seul le complément qui a créé les paramètres de complément de contenu et du volet Office peut y accéder à partir du document où ils sont enregistrés.<br/><br/>**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.|-Mot<br>-Excel<br>-Powerpoint<br/><br/> **Remarque :** les compléments du volet Office pour Project 2013 ne prennent pas en charge l’API **Settings** pour le stockage de l’état ou des paramètres du complément. Toutefois, pour les compléments s’exécutant dans Project (ainsi que dans d’autres applications clientes Office), vous pouvez utiliser des techniques telles que les cookies de navigateur ou le stockage web. Pour plus d’informations sur ces techniques, reportez-vous à l’exemple de code [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|mail|Boîte aux lettres de serveur Exchange de l’utilisateur où le complément est installé. Étant donné que ces paramètres sont stockés dans la boîte aux lettres du serveur de l’utilisateur, ils peuvent être « itinérants » avec l’utilisateur et sont disponibles pour le complément lorsqu’il s’exécute dans le contexte d’une application cliente Office ou d’un navigateur pris en charge accédant à la boîte aux lettres de cet utilisateur.<br/><br/> Seul le complément qui a créé les paramètres d’itinérance du complément Outlook peut y accéder, et uniquement dans la boîte aux lettres où le complément est installé.|Outlook|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|mail|The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.|Outlook|
|[CustomXmlParts](/javascript/api/office/office.customxmlparts)|volet Office|The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.<br/><br/>**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.|- Word (à l’aide de l’API commune JavaScript Office)<br>- Excel (à l’aide de l’API JavaScript Excel spécifique à l’application)|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>Données de paramètres gérées en mémoire à l’exécution

> [!NOTE]
> Les deux sections suivantes abordent les paramètres dans le contexte de l’API JavaScript courante pour Office. L’API JavaScript Excel spécifique à l’application permet également d’accéder aux paramètres personnalisés. Les API Excel et les modes de programmation sont légèrement différents. Pour plus d’informations, reportez-vous à l’article sur l’objet [SettingCollection pour Excel](/javascript/api/excel/excel.settingcollection).

En interne, les données du conteneur de propriétés accessibles avec les `Settings`objets , `CustomProperties`ou `RoamingSettings` sont stockées sous la forme d’un objet JSON (JavaScript Object Notation) sérialisé qui contient des paires nom/valeur. Le nom (clé) de chaque valeur doit être , `string`et la valeur stockée peut être javaScript `string`, `number`, `date`ou `object`, mais pas une **fonction**.

Cet exemple de structure de conteneur des propriétés contient trois valeurs de type **string** (chaîne) définies, nommées `firstName`, `location` et `defaultView`.

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

Après avoir enregistré le conteneur des propriétés de paramètres durant la session de complément précédente, vous pouvez le charger pendant ou après l’initialisation du complément, durant la session actuelle du complément. Pendant la session, les paramètres sont gérés entièrement en mémoire à l’aide `get`des méthodes , `set`et `remove` de l’objet qui correspond au type de paramètres que vous créez (**Paramètres**, **CustomProperties** ou **RoamingSettings**).

> [!IMPORTANT]
> Pour conserver les ajouts, mises à jour ou suppressions effectués pendant la session active du complément dans l’emplacement de stockage, vous devez appeler la `saveAsync` méthode de l’objet correspondant utilisé pour utiliser ce type de paramètres. Les `get`méthodes , `set`et `remove` fonctionnent uniquement sur la copie en mémoire du conteneur de propriétés des paramètres. Si votre complément est fermé sans appeler `saveAsync`, toutes les modifications apportées aux paramètres au cours de cette session seront perdues.

## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>Enregistrement de l’état et des paramètres d’un complément par document pour les compléments de contenu et du volet Office

Pour conserver l’état ou les paramètres personnalisés d’un complément de contenu ou du volet Office pour Word, Excel ou PowerPoint, utilisez l’objet [Settings](/javascript/api/office/office.settings) et ses méthodes. Le conteneur de propriétés créé avec les méthodes de l’objet `Settings` est disponible uniquement pour l’instance du complément de contenu ou du volet Office qui l’a créé, et uniquement à partir du document dans lequel il est enregistré.

L’objet `Settings` est automatiquement chargé dans le cadre de l’objet [Document](/javascript/api/office/office.document) et est disponible lorsque le complément de volet Office ou de contenu est activé. Une fois l’objet `Document` instancié, vous pouvez accéder à l’objet `Settings` avec la propriété [settings](/javascript/api/office/office.document#office-office-document-settings-member) de l’objet `Document` . Pendant la durée de vie de la session, vous pouvez simplement utiliser les `Settings.get`méthodes , `Settings.set`et `Settings.remove` pour lire, écrire ou supprimer les paramètres persistants et l’état du complément de la copie en mémoire du conteneur de propriétés.

Étant donné que les méthodes de définition (set) et de suppression (remove) fonctionnent uniquement par rapport à la copie en mémoire du conteneur des propriétés de paramètres, pour enregistrer de nouveaux paramètres ou des paramètres modifiés dans le document auquel le complément est associé, vous devez appeler la méthode [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)).

### <a name="creating-or-updating-a-setting-value"></a>Création ou mise à jour d’une valeur de paramètre

The following code example shows how to use the [Settings.set](/javascript/api/office/office.settings#office-office-settings-set-member(1)) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.

```js
Office.context.document.settings.set('themeColor', 'green');
```

 Le paramètre avec le nom spécifié est créé s’il n’existe pas déjà ou sa valeur est mise à jour s’il existe. Utilisez la `Settings.saveAsync` méthode pour conserver les paramètres nouveaux ou mis à jour dans le document.

### <a name="getting-the-value-of-a-setting"></a>Obtention de la valeur d’un paramètre

L’exemple suivant illustre comment utiliser la méthode [Settings.get](/javascript/api/office/office.settings#office-office-settings-get-member(1)) pour obtenir la valeur d’un paramètre nommé « themeColor ». Le seul paramètre de la `get` méthode est le _nom_ du paramètre qui respecte la casse.

```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

 La `get` méthode retourne la valeur précédemment enregistrée pour le _nom_ du paramètre qui a été passé. Si le paramètre n’existe pas, la méthode retourne **null**.

### <a name="removing-a-setting"></a>Suppression d’un paramètre

L’exemple suivant illustre comment utiliser la méthode [Settings.remove](/javascript/api/office/office.settings#office-office-settings-remove-member(1)) pour supprimer un paramètre portant le nom « themeColor ». Le seul paramètre de la `remove` méthode est le _nom_ du paramètre qui respecte la casse.

```js
Office.context.document.settings.remove('themeColor');
```

Rien ne se produit si le paramètre n’existe pas. Utilisez la `Settings.saveAsync` méthode pour conserver la suppression du paramètre du document.

### <a name="saving-your-settings"></a>Enregistrement de vos paramètres

Pour enregistrer les ajouts, modifications ou suppressions que votre complément a effectués sur la copie en mémoire du conteneur de propriétés des paramètres pendant la session en cours, vous devez appeler la méthode [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)) pour les stocker dans le document. Le seul paramètre de la `saveAsync` méthode est _callback_, qui est une fonction de rappel avec un seul paramètre.

```js
Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        write('Settings saved.');
    }
});
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

La fonction anonyme transmise à la `saveAsync` méthode en tant que paramètre _de rappel_ est exécutée lorsque l’opération est terminée. Le paramètre _asyncResult_ du rappel permet d’accéder à un `AsyncResult` objet qui contient l’état de l’opération. Dans l’exemple, la fonction vérifie la `AsyncResult.status` propriété pour voir si l’opération d’enregistrement a réussi ou échoué, puis affiche le résultat dans la page du complément.

## <a name="how-to-save-custom-xml-to-the-document"></a>Enregistrement du XML personnalisé dans le document

> [!NOTE]
> Cette section décrit les parties XML personnalisées dans le contexte de l’API JavaScript courante pour Office qui est prise en charge dans Word. L’API JavaScript Excel spécifique à l’application permet également d’accéder aux parties XML personnalisées. Les API Excel et les modes de programmation sont légèrement différents. Pour plus d’informations, reportez-vous à l’article sur l’objet [CustomXmlPart pour Excel](/javascript/api/excel/excel.customxmlpart).

Il existe une option de stockage supplémentaire lorsque vous devez stocker des informations qui dépassent les limites de taille des paramètres du document ou qui ont un caractère structuré. Vous pouvez conserver le balisage XML personnalisé dans un complément de volet Office pour Word (et pour Excel, mais reportez-vous à la remarque en haut de cette section). Dans Word, utilisez l’objet [CustomXmlPart](/javascript/api/office/office.customxmlpart) et ses méthodes (rappel : pour Excel, consultez la note précédente). Le code suivant crée une partie XML personnalisée, puis affiche son ID et son contenu dans des balises div sur la page. Un attribut `xmlns` doit figurer dans la chaîne XML.

```js
function createCustomXmlPart() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            $("#xml-id").text("Your new XML part's ID: " + asyncResult.value.id);
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

Pour récupérer une partie XML personnalisée, vous utilisez la méthode [getByIdAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbyidasync-member(1)), mais l’identifiant correspond à un GUID généré lorsque la partie XML est créée. Vous ne pouvez donc pas connaître l’identifiant lors du codage. Pour cette raison, il est recommandé de stocker immédiatement l’identifiant de la partie XML en tant que paramètre et de lui donner une clé facilement mémorisable lorsque vous créez une partie XML. L’exemple de méthode suivant montre comment procéder. (Mais consultez les sections précédentes de cet article pour plus d’informations et les meilleures pratiques lors de l’utilisation de paramètres personnalisés.)

 ```js
function createCustomXmlPartAndStoreId() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            Office.context.document.settings.set('ReviewersID', asyncResult.id);
            Office.context.document.settings.saveAsync();
        }
    );
}
```

Le code suivant montre comment récupérer la partie XML en obtenant d’abord son identifiant partir d’un paramètre.

 ```js
function getReviewers() {
    const reviewersXmlId = Office.context.document.settings.get('ReviewersID');
    Office.context.document.customXmlParts.getByIdAsync(reviewersXmlId,
        (asyncResult) => {
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

## <a name="how-to-save-settings-in-an-outlook-add-in"></a>Comment enregistrer des paramètres dans un complément Outlook

Pour plus d’informations sur l’enregistrement des paramètres dans un complément Outlook, voir [Gérer l’état et les paramètres d’un complément Outlook](../outlook/manage-state-and-settings-outlook.md).

## <a name="see-also"></a>Voir aussi

- [Compréhension de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
- [Compléments Outlook](../outlook/outlook-add-ins-overview.md)
- [Gérer l’état et les paramètres d’un complément Outlook](../outlook/manage-state-and-settings-outlook.md)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
