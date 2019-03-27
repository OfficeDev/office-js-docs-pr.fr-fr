# <a name="whats-changed-in-the-javascript-api-for-office"></a>Nouveautés de l’API JavaScript pour Office

Afin d’étendre la fonctionnalité de vos Compléments Office, des objets, méthodes, propriétés, événements et énumérations sont régulièrement ajoutés et mis à jour dans l’API JavaScript pour Office. Utilisez les liens ci-dessous pour afficher les membres de l’API qui ont été ajoutés ou mis à jour.

Pour développer des compléments utilisant les nouveaux membres de l’API, vous devez [mettre à jour l’API JavaScript pour les fichiers de l’API JavaScript pour Office dans votre projet](/office/dev/add-ins/develop/update-your-javascript-api-for-office-and-manifest-schema-version).

Pour visualiser tous les membres de l’API, y compris ceux qui sont identiques par rapport aux versions précédentes, voir [API JavaScript pour Office](javascript-api-for-office.md).

## <a name="new-and-updated-apis"></a>API ajoutées et mises à jour

### <a name="new-and-updated-objects"></a>Objets ajoutés et mis à jour

|**Objet**|**Description**|**Version ajoutée ou mise à jour**|
|:-----|:-----|:-----|
|`Item`|Mis à jour et ajouté pour :<br><ul><li><p>Méthodes `getSelectedDataAsync` et `setSelectedDataAsync` pour prendre en charge l’obtention de la sélection de l’utilisateur et son remplacement dans l’objet et le corps d’un message ou d’un rendez-vous.</p></li><li><p>Méthodes `displayReplyAllForm` et `displayReplyForm` pour prendre en charge l’ajout d’une pièce jointe au formulaire de réponse d’un rendez-vous.</p></li></ul>|Boîte aux lettres 1.2|
|`Item`|Mis à jour pour inclure des méthodes et des champs utiles à la création de compléments Outlook en mode composition. |1.1|
|`Binding`|Mis à jour pour prendre en charge la liaison de tableau dans les compléments de contenu pour Access.|1.1|
|`Bindings`|Mis à jour pour prendre en charge la liaison de tableau dans les compléments de contenu pour Access.|1.1|
|`Body`|Ajouté pour permettre la création et la modification du corps d’un message ou d’un rendez-vous dans les compléments Outlook en mode composition.|1.1|
|`Document`|Mises à jour et ajouts pour les éléments suivants : <ul><li><p>Prendre en charge les propriétés <a href="/javascript/api/office/office.document" target="_blank">mode</a>, <a href="/javascript/api/office/office.document#settings" target="_blank">settings</a> et <a href="/javascript/api/office/office.document" target="_blank">url</a> dans les compléments de contenu pour Access.</p></li><li><p>Obtenir le document au format PDF à l’aide de la méthode <a href="/javascript/api/office/office.document#getfileasync-filetype--options--callback-" target="_blank">getFileAsync</a> dans les compléments pour PowerPoint et Word.</p></li><li><p>Obtenir les propriétés de fichier à l’aide de la méthode <a href="/javascript/api/office/office.document#getfilepropertiesasync-options--callback-" target="_blank">getFileProperties</a> dans les compléments pour Excel, PowerPoint et Word.</p></li><li><p>Accéder aux emplacements et aux objets au sein du document à l’aide de la méthode <a href="/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-" target="_blank">goToByIdAsync</a> dans les compléments pour Excel et PowerPoint.</p></li><li><p>Obtenir l’ID, le titre et l’index des diapositives sélectionnées à l’aide de la méthode <a href="/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-" target="_blank">getSelectedDataAsync</a> (lorsque vous spécifiez la nouvelle énumération <span class="keyword">Office.CoercionType.SlideRange</span><a href="/javascript/api/office/office.coerciontype" target="_blank">coercionType</a>) dans les compléments pour PowerPoint.</p></li></ul>|1.1|
|`Location`|Ajouté pour permettre la définition de l’emplacement d’un rendez-vous dans les compléments Outlook en mode composition.|1.1|
|`Office`|Mise à jour de la méthode Select pour prendre en charge l’obtention des liaisons dans les compléments de contenu pour Access.|1.1|
|`Recipients`|Ajouté pour permettre l’obtention et la définition des destinataires d’un message ou d’un rendez-vous en mode composition.|1.1|
|`Settings`|Mis à jour pour prendre en charge la création de paramètres personnalisés dans les compléments de contenu pour Access.|1.1|
|`Subject`|Ajouté pour permettre l’obtention et la définition de l’objet d’un message ou d’un rendez-vous dans les compléments Outlook en mode composition.|1.1|
|`Time`|Ajouté pour permettre l’obtention et la définition de l’heure de début et de fin d’un rendez-vous dans les compléments Outlook en mode composition.|1.1|

### <a name="new-and-updated-enumerations"></a>Énumérations ajoutées et énumérations mises à jour

|**Objet**|**Description**|**Version**|
|:-----|:-----|:-----|
|`ActiveView`|Spécifie l’état de l’affichage dynamique du document, par exemple, si l’utilisateur peut modifier le document.Ajouté pour permettre aux compléments pour PowerPoint de déterminer si un utilisateur visualise une présentation ( **Diaporama**) ou modifie des diapositives. |1.1|
|`CoercionType`|Mis à jour avec **Office.CoercionType.SlideRange** pour permettre la prise en charge de l’obtention des diapositives sélectionnées à l’aide de la méthode **getSelectedDataAsync** dans les compléments pour PowerPoint.|1.1|
|`EventType`|Mis à jour pour inclure le nouvel événement ActiveViewChanged.|1.1|
|`FileType`|Mis à jour pour spécifier la sortie au format PDF.|1.1|
|`GoToType`|Ajouté pour spécifier l’emplacement ou l’objet auquel accéder dans le document.|1.1|

