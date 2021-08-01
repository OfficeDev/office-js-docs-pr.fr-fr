---
title: Conservation de l’état et des paramètres des compléments
description: Apprenez à faire persister des données dans Office applications web de recherche de contenu s’exécutant dans l’environnement sans état d’un contrôle de navigateur.
ms.date: 03/23/2021
localization_priority: Normal
ms.openlocfilehash: a1730826fafb840bc6ae69c5c1e95c54ccc0f2f7
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671519"
---
# <a name="persisting-add-in-state-and-settings"></a>Conservation de l’état et des paramètres des compléments

[!include[information about the common API](../includes/alert-common-api-info.md)]

Les compléments Office sont essentiellement des applications web exécutées dans l’environnement sans état d’un contrôle de navigateur. En conséquence, votre complément devra peut-être faire persister les données pour assurer la continuité de certaines opérations ou fonctionnalités entre les sessions d’utilisation du complément. Par exemple, votre complément peut disposer de paramètres personnalisés ou d’autres valeurs dont il a besoin pour l’enregistrement et le rechargement à la prochaine initialisation, tels que l’affichage préféré d’un utilisateur ou l’emplacement par défaut. Pour ce faire, vous pouvez procéder comme suit :

- Utilisez les membres de l’API JavaScript Office qui stockent les données sous l’une des formes :
  - Paires nom/valeur dans un conteneur de propriétés stocké dans un emplacement qui dépend du type de complément.
  - Éléments XML personnalisés stockés dans le document.

- Utilisez des techniques fournies par le contrôle de navigateur sous-jacent : les cookies de navigateur ou le stockage web HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) ou [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).
    > [!NOTE]
    > L’utilisateur peut bloquer les techniques de stockage basées sur le navigateur en fonction des paramètres qu’il choisit.

Cet article se concentre sur l’utilisation de l Office API JavaScript pour rendre persistant l’état du module dans le document actuel. Si vous devez rendre persistant l’état entre les documents, par exemple suivre les préférences des utilisateurs dans tous les documents qu’ils ouvrent, vous devez utiliser une approche différente. Par exemple, vous [](sso-in-office-add-ins.md#using-the-sso-token-as-an-identity) pouvez utiliser l' utilisateur unique pour obtenir l’identité de l’utilisateur, puis enregistrer l’ID utilisateur et ses paramètres dans une base de données en ligne.

## <a name="persisting-add-in-state-and-settings-with-the-office-javascript-api"></a>Persistance de l’état et des paramètres d’un Office avec l’API JavaScript

L Office API JavaScript fournit les objets [Paramètres,](/javascript/api/office/office.settings) [RoamingSettings](/javascript/api/outlook/office.roamingsettings)et [CustomProperties](/javascript/api/outlook/office.customproperties) pour enregistrer l’état du module dans les sessions, comme décrit dans le tableau suivant. Dans tous les cas, les valeurs de paramètre enregistrées sont associées à l’[ID](../reference/manifest/id.md) du complément qui les a créées.

|**Objet**|**Type de complément**|**Emplacement de stockage**|**Office prise en charge des applications**|
|:-----|:-----|:-----|:-----|
|[Paramètres](/javascript/api/office/office.settings)|Contenu et volet de tâches|Document, feuille de calcul ou présentation qu’utilise le complément. Seul le complément qui a créé les paramètres de complément de contenu et du volet Office peut y accéder à partir du document où ils sont enregistrés.<br/><br/>**Remarque importante :** ne stockez pas de mots de passe ou autres informations d’identification personnelle (PII) avec l’objet **Settings**. Les données enregistrées ne sont pas visibles par les utilisateurs finals. Toutefois, elles sont stockées en tant que partie du document, qui est accessible en lisant directement le format de fichier. Vous devez limiter l’utilisation de PII de votre complément et stocker ces informations requises par votre complément uniquement sur le serveur qui l’héberge en tant que ressource sécurisée par l’utilisateur.|Word, Excel ou PowerPoint<br/><br/> **Remarque :** les compléments du volet Office pour Project 2013 ne prennent pas en charge l’API **Settings** pour le stockage de l’état ou des paramètres du complément. Toutefois, pour les applications Project (ainsi que d’autres applications clientes Office), vous pouvez utiliser des techniques telles que les cookies de navigateur ou le stockage web. Pour plus d’informations sur ces techniques, reportez-vous à l’exemple de code [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|Outlook|Boîte aux lettres de serveur Exchange de l’utilisateur où le complément est installé. Étant donné que ces paramètres sont stockés dans la boîte aux lettres du serveur de l’utilisateur, ils peuvent être « itinérants » avec l’utilisateur et sont disponibles pour le module lorsqu’il est en cours d’exécution dans le contexte d’une application cliente ou d’un navigateur Office pris en charge accédant à la boîte aux lettres de cet utilisateur.<br/><br/> Seul le complément qui a créé les paramètres d’itinérance du complément Outlook peut y accéder, et uniquement dans la boîte aux lettres où le complément est installé.|Outlook|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|Outlook|Élément de message, de rendez-vous ou de demande de réunion qu’utilise le complément. Seul le complément qui a créé les propriétés personnalisées d’élément de complément Outlook peut y accéder, et uniquement dans l’élément où elles sont enregistrées.|Outlook|
|[CustomXmlParts](/javascript/api/office/office.customxmlparts)|volet Office|Document, feuille de calcul ou présentation que le complément utilise. Les paramètres de complément de volet Office sont disponibles pour le complément qui les a créés dans le document dans lequel ils sont enregistrés.<br/><br/>**Important :** ne stockez pas de mot de passe ni d’autres informations d’identification personnelle dans une partie XML personnalisée. Les données enregistrées ne sont pas visibles par les utilisateurs finals. Toutefois, elles sont stockées en tant que partie du document, qui est accessible en lisant directement le format de fichier. Vous devez limiter l’utilisation des informations d’identification personnelle de votre complément et stocker ces informations requises par votre complément uniquement sur le serveur qui l’héberge en tant que ressource sécurisée par l’utilisateur.|Word (à l’aide Office’API commune JavaScript) Excel (à l’aide de l’API JavaScript Excel l’application)|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>Données de paramètres gérées en mémoire à l’exécution

> [!NOTE]
> Les deux sections suivantes abordent les paramètres dans le contexte de l’API JavaScript courante pour Office. L’api JavaScript Excel l’application fournit également l’accès aux paramètres personnalisés. Les API Excel et les modes de programmation sont légèrement différents. Pour plus d’informations, reportez-vous à l’article sur l’objet [SettingCollection pour Excel](/javascript/api/excel/excel.settingcollection).

En interne, les données du sac de propriétés accessibles avec le ou les objets sont stockées sous la forme d’un objet JSON (JavaScript Object Notation) sérialisé qui contient des `Settings` `CustomProperties` `RoamingSettings` paires nom/valeur. Le nom (clé) de chaque valeur doit être un , et la valeur stockée peut être un JavaScript , ou , mais pas `string` `string` une `number` `date` `object` **fonction**.

Cet exemple de structure de conteneur des propriétés contient trois valeurs de type **string** (chaîne) définies, nommées `firstName`, `location` et `defaultView`.

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

Après avoir enregistré le conteneur des propriétés de paramètres durant la session de complément précédente, vous pouvez le charger pendant ou après l’initialisation du complément, durant la session actuelle du complément. Pendant la session, les paramètres sont gérés entièrement en mémoire à l’aide de l’objet , et des méthodes de l’objet qui correspond au type de paramètres que vous créez `get` `set` ( `remove` **Paramètres,** **CustomProperties** ou **RoamingSettings**).

> [!IMPORTANT]
> Pour rendre persistants les ajouts, mises à jour ou suppressions effectués au cours de la session actuelle du complément vers l’emplacement de stockage, vous devez appeler la méthode de l’objet correspondant utilisé pour travailler avec ce type de `saveAsync` paramètres. Les méthodes et les paramètres fonctionnent uniquement sur la copie en mémoire du sac des `get` `set` `remove` propriétés des paramètres. Si votre add-in est fermé sans appel, toutes les modifications apportées aux paramètres au cours de `saveAsync` cette session seront perdues.

## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>Enregistrement de l’état et des paramètres d’un complément par document pour les compléments de contenu et du volet Office

Pour conserver l’état ou les paramètres personnalisés d’un complément de contenu ou du volet Office pour Word, Excel ou PowerPoint, utilisez l’objet [Settings](/javascript/api/office/office.settings) et ses méthodes. Le sac de propriétés créé avec les méthodes de l’objet est disponible uniquement pour l’instance du contenu ou du module de tâches qui l’a créé, et uniquement à partir du document dans lequel il est `Settings` enregistré.

L’objet est automatiquement chargé dans le cadre de l’objet Document et est disponible lorsque le volet Des tâches ou le module de contenu `Settings` est activé. [](/javascript/api/office/office.document) Une fois l’objet ins instantié, vous pouvez accéder à l’objet avec la propriété `Document` `Settings` [settings](/javascript/api/office/office.document#settings) de `Document` l’objet. Pendant toute la durée de vie de la session, vous pouvez simplement utiliser le , et les méthodes pour lire, écrire ou supprimer les paramètres persistants et l’état du add-in de la copie en mémoire du sac des `Settings.get` `Settings.set` `Settings.remove` propriétés.

Étant donné que les méthodes de définition (set) et de suppression (remove) fonctionnent uniquement par rapport à la copie en mémoire du conteneur des propriétés de paramètres, pour enregistrer de nouveaux paramètres ou des paramètres modifiés dans le document auquel le complément est associé, vous devez appeler la méthode [Settings.saveAsync](/javascript/api/office/office.settings#saveAsync_options__callback_).

### <a name="creating-or-updating-a-setting-value"></a>Création ou mise à jour d’une valeur de paramètre

L’exemple de code suivant montre comment utiliser la méthode [Settings.set](/javascript/api/office/office.settings#set_name__value_) pour créer un paramètre appelé `'themeColor'` avec la valeur `'green'`. Le premier paramètre de la méthode set est le _name_ (ID) respectant la casse du paramètre à définir ou à créer. Le second paramètre est la _value_ du paramètre.

```js
Office.context.document.settings.set('themeColor', 'green');
```

 Le paramètre avec le nom spécifié est créé s’il n’existe pas déjà ou sa valeur est mise à jour s’il existe. Utilisez la méthode pour rendre persistants les paramètres nouveaux ou mis à `Settings.saveAsync` jour dans le document.

### <a name="getting-the-value-of-a-setting"></a>Obtention de la valeur d’un paramètre

L’exemple suivant illustre comment utiliser la méthode [Settings.get](/javascript/api/office/office.settings#get_name_) pour obtenir la valeur d’un paramètre nommé « themeColor ». Le seul paramètre de la méthode est le nom du paramètre sensible `get` _à la_ cas.

```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

 La méthode renvoie la valeur précédemment enregistrée pour le nom du paramètre `get` qui a été transmis.  Si le paramètre n’existe pas, la méthode retourne **null**.

### <a name="removing-a-setting"></a>Suppression d’un paramètre

L’exemple suivant illustre comment utiliser la méthode [Settings.remove](/javascript/api/office/office.settings#remove_name_) pour supprimer un paramètre portant le nom « themeColor ». Le seul paramètre de la méthode est le nom du paramètre sensible `remove` _à la_ cas.

```js
Office.context.document.settings.remove('themeColor');
```

Rien ne se produit si le paramètre n’existe pas. Utilisez la `Settings.saveAsync` méthode pour rendre persistant la suppression du paramètre du document.

### <a name="saving-your-settings"></a>Enregistrement de vos paramètres

Pour enregistrer les ajouts, modifications ou suppressions que votre complément a effectués sur la copie en mémoire du conteneur de propriétés des paramètres pendant la session en cours, vous devez appeler la méthode [Settings.saveAsync](/javascript/api/office/office.settings#saveAsync_options__callback_) pour les stocker dans le document. Le seul paramètre de la méthode est le rappel, qui est une fonction `saveAsync` de rappel avec un seul paramètre. 

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

La fonction anonyme transmise dans la méthode en tant que paramètre de rappel est `saveAsync` exécutée lorsque l’opération  est terminée. Le _paramètre asyncResult_ du rappel permet d’accéder à un objet qui contient `AsyncResult` l’état de l’opération. Dans l’exemple, la fonction vérifie la propriété pour voir si l’opération d’enregistrer a réussi ou échoué, puis affiche le résultat dans la page du `AsyncResult.status` module.

## <a name="how-to-save-custom-xml-to-the-document"></a>Enregistrement du XML personnalisé dans le document

> [!NOTE]
> Cette section décrit les parties XML personnalisées dans le contexte de l’API JavaScript courante pour Office qui est prise en charge dans Word. L’api JavaScript Excel l’application fournit également l’accès aux parties XML personnalisées. Les API Excel et les modes de programmation sont légèrement différents. Pour plus d’informations, reportez-vous à l’article sur l’objet [CustomXmlPart pour Excel](/javascript/api/excel/excel.customxmlpart).

Il existe une option de stockage supplémentaire lorsque vous devez stocker des informations qui dépassent les limites de taille de la Paramètres document ou qui ont un caractère structuré. Vous pouvez conserver le balisage XML personnalisé dans un complément de volet Office pour Word (et pour Excel, mais reportez-vous à la remarque en haut de cette section). Dans Word, utilisez l’objet [CustomXmlPart](/javascript/api/office/office.customxmlpart) et ses méthodes (rappel : pour Excel, consultez la note précédente). Le code suivant crée une partie XML personnalisée, puis affiche son ID et son contenu dans des balises div sur la page. Un attribut `xmlns` doit figurer dans la chaîne XML.

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

Pour récupérer une partie XML personnalisée, vous utilisez la méthode [getByIdAsync](/javascript/api/office/office.customxmlparts#getByIdAsync_id__options__callback_), mais l’identifiant correspond à un GUID généré lorsque la partie XML est créée. Vous ne pouvez donc pas connaître l’identifiant lors du codage. Pour cette raison, il est recommandé de stocker immédiatement l’identifiant de la partie XML en tant que paramètre et de lui donner une clé facilement mémorisable lorsque vous créez une partie XML. L’exemple de méthode suivant montre comment procéder. (Mais consultez les sections précédentes de cet article pour plus d’informations et les meilleures pratiques concernant l’utilisation de paramètres personnalisés.)

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

## <a name="how-to-save-settings-in-an-outlook-add-in"></a>Comment enregistrer des paramètres dans un Outlook de configuration

Pour plus d’informations sur l’enregistrer dans un Outlook, voir Gérer l’état et [les paramètres d’un Outlook.](../outlook/manage-state-and-settings-outlook.md)

## <a name="see-also"></a>Voir aussi

- [Compréhension de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
- [Compléments Outlook](../outlook/outlook-add-ins-overview.md)
- [Gérer l’état et les paramètres d’un Outlook de gestion](../outlook/manage-state-and-settings-outlook.md)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
