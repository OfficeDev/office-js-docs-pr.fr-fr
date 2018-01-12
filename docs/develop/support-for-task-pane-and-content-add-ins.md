
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Prise en charge de l’interface API JavaScript pour Office pour les compléments de contenu et du volet Office dans Office 2013


Vous pouvez utiliser l’[API JavaScript pour Office](../../reference/javascript-api-for-office.md) pour créer un complément du volet Office ou de contenu pour les applications hôtes d’Office 2013. Les méthodes et les objets pris en charge par les compléments du volet Office et de contenu sont classés de la manière suivante :


1. **Objets communs partagés avec d’autres compléments Office.** Parmi ces objets figurent [Office](../../reference/shared/office.md), [Context](../../reference/shared/office.context.md) et [AsyncResult](../../reference/shared/asyncresult.md). L’objet **Office** est l’objet racine de l’interface API JavaScript pour Office. L’objet **Context** représente l’environnement d’exécution du complément. **Office** et **Context** sont les objets fondamentaux pour tout complément Office. L’objet **AsyncResult** représente les résultats d’une opération asynchrone, comme les données renvoyées vers la méthode **getSelectedDataAsync**, qui lit les éléments sélectionnés par un utilisateur dans un document.
    
2.  **Objet Document** Une grand partie de l’API disponible pour les compléments de contenu ou du volet Office est exposée via les méthodes, les propriétés et les événements de l’objet [Document](../../reference/shared/document.md). Un complément de contenu ou du volet Office peut utiliser la propriété [Office.context.document](../../reference/shared/office.context.document.md) pour accéder à l’objet **Document**, et accéder par ce biais aux membres clés de l’API pour utiliser des données dans des documents, telles que les objets [Bindings](../../reference/shared/bindings.bindings.md) et [CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md), et les méthodes [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md), [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) et [getFileAsync](../../reference/shared/document.getfileasync.md). L’objet **Document**  fournit également la propriété [mode](../../reference/shared/document.mode.md) permettant de déterminer si un document est en mode lecture seule ou modification, la propriété [url](../../reference/shared/document.url.md) pour obtenir l’URL du document actuel et accéder à l’objet [Settings](../../reference/shared/settings.md). L’objet **Document** prend également en charge l’ajout de gestionnaires d’événements pour l’événement [SelectionChanged](../../reference/shared/document.selectionchanged.event.md), afin de vous permettre de détecter quand un utilisateur modifie sa sélection dans le document.
    
   Un complément de contenu ou du volet Office peut accéder à l’objet **Document** uniquement après le chargement de l’environnement d’exécution et du DOM, généralement dans le gestionnaire d’événements pour l’événement [Office.initialize](../../reference/shared/office.initialize.md). Pour plus d’informations sur le flux d’événements lors de l’initialisation d’un complément et sur la vérification du chargement correct du DOM et de l’environnement d’exécution, voir la page relative au [chargement du DOM et de l’environnement d’exécution](../../docs/develop/loading-the-dom-and-runtime-environment.md).
    
3.  **Objets pour l’utilisation de fonctionnalités spécifiques.** Pour travailler avec des fonctionnalités spécifiques de l’API, utilisez les méthodes et les objets suivants :
    
    - Les objets [CustomXmlParts](../../reference/shared/bindings.bindings.md), [CustomXmlPart](../../reference/shared/binding.md) et les objets associés pour créer et manipuler des parties XML personnalisées dans des documents Word.
    
    - Les objets [File](../../reference/shared/customxmlparts.customxmlparts.md) et [Slice](../../reference/shared/customxmlpart.customxmlpart.md) pour créer une copie de l’intégralité du document, le diviser en blocs ou en « sections », puis lire ou transmettre les données dans ces sections.
    
    - Les objets [File](../../reference/shared/file.md) et [Slice](../../reference/shared/slice.md) pour créer une copie de l’intégralité du document, le diviser en blocs ou en « sections », puis lire ou transmettre les données dans ces sections.
    
    - [Important](../../reference/shared/settings.md)  Certains des membres de l’API ne sont pas pris en charge dans toutes les applications Office pouvant héberger des compléments de contenu et du volet Office. Pour déterminer les membres pris en charge, voir les ressources suivantes :
    

 >**Important**  Certains des membres de l’API ne sont pas pris en charge dans toutes les applications Office pouvant héberger des compléments de contenu et du volet Office.

Pour consulter un résumé de la prise en charge de l’interface API JavaScript pour Office dans les applications hôtes d’Office, voir [Présentation de l’interface API Javascript pour Office](../../docs/develop/understanding-the-javascript-api-for-office.md).


## <a name="reading-and-writing-to-an-active-selection"></a>Lecture et écriture dans une sélection active

Vous pouvez lire ou écrire dans la sélection en cours de l’utilisateur dans un document, une feuille de calcul ou une présentation. Selon l’application hôte de votre complément, vous pouvez spécifier le type de structure de données à lire ou à écrire en tant que paramètre dans les méthodes [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) et [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) de l’objet [Document](../../reference/shared/document.md). Par exemple, vous pouvez indiquer n’importe quel type de données (HTML, données tabulaires, Office Open XML ou texte) pour Word, des données texte et tabulaires pour Excel et des données texte pour PowerPoint et Project. Vous pouvez également créer des gestionnaires d’événements pour détecter les modifications apportées à la sélection de l’utilisateur. L’exemple suivant obtient des données à partir de la sélection en tant que données texte à l’aide de la méthode **getSelectedDataAsync**.


```js
Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        }
        else {
            write('Selected data: ' + asyncResult.value);
        }
    });

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

Pour plus d’informations et d’exemples, voir l’article concernant la [lecture et l’écriture de données dans la sélection active d’un document ou d’une feuille de calcul](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## <a name="binding-to-a-region-in-a-document-or-spreadsheet"></a>Liaison à une région d’un document ou d’une feuille de calcul

Vous pouvez utiliser les méthodes **getSelectedDataAsync** et **setSelectedDataAsync** pour lire ou écrire la sélection *en cours* de l’utilisateur dans un document, une feuille de calcul ou une présentation.

Vous pouvez ajouter une liaison à l’aide des méthodes [addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md), [addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md) ou [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) de l’objet [Bindings](../../reference/shared/bindings.bindings.md).

Pour plus d’informations et d’exemples, voir **Liaisons de régions dans un document ou une feuille de calcul**.



```js
Office.context.document.bindings.addFromSelectionAsync(
    Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' +
            asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Pour plus d’informations et d’exemples, voir l’article [Liaisons de régions dans un document ou une feuille de calcul](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="getting-entire-documents"></a>Obtention de documents entiers

Si votre complément du volet Office s’exécute dans PowerPoint ou Word, vous pouvez utiliser les méthodes [Document.getFileAsync](../../reference/shared/document.getfileasync.md), [File.getSliceAsync](../../reference/shared/file.getsliceasync.md) et [File.closeAsync](../../reference/shared/file.closeasync.md) pour obtenir la totalité d’une présentation ou d’un document.

Lorsque vous appelez **Document.getFileAsync**, vous obtenez une copie du document dans un objet [File](../../reference/shared/file.md). L’objet **File** donne accès au document en « blocs » représenté sous la forme d’objets [Slice](../../reference/shared/document.md). Lorsque vous appelez **getFileAsync**, vous pouvez spécifier le type de fichier (texte ou format Open Office XML compressé) et la taille des secteurs (jusqu’à 4 Mo). Pour accéder au contenu de l’objet **File**, appelez **File.getSliceAsync** qui renvoie les données brutes dans la propriété [Slice.data](../../reference/shared/slice.data.md). Si vous avez spécifié un format compressé, vous obtiendrez les données du fichier sous la forme d’un tableau d’octets. Si vous transférez le fichier à un service web, vous pouvez transformer les données brutes compressées dans une chaîne codée en Base64 avant l’envoi. Enfin, lorsque vous avez obtenu les sections du fichier, utilisez la méthode **File.closeAsync** pour fermer le document.

Pour plus d’informations, voir l’article relatif à la [façon d’obtenir l’intégralité d’un document à partir d’un complément pour PowerPoint ou Word](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md). 


## <a name="reading-and-writing-custom-xml-parts-of-a-word-document"></a>Lecture et écriture des parties XML personnalisées d’un document Word

Grâce aux contrôles de contenu et au format de fichier Office Open XML, vous pouvez ajouter des parties XML personnalisées à un document Word et lier des éléments dans les parties XML aux contrôles de contenu de ce document. Lorsque vous ouvrez le document, Word lit et remplit automatiquement les contrôles de contenu liés avec les données des parties XML personnalisées. Les utilisateurs peuvent également écrire des données dans les contrôles de contenu. Lorsqu’ils enregistrent le document, les données des contrôles sont alors enregistrées dans les parties XML liées. Si votre complément du volet Office s’exécute dans Word, vous pouvez utiliser la propriété [Document.customXmlParts](../../reference/shared/document.customxmlparts.md), ainsi que les objets [CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md), [CustomXmlPart](../../reference/shared/customxmlpart.customxmlpart.md) et [CustomXmlNode](../../reference/shared/customxmlnode.customxmlnode.md) pour lire et écrire des données de manière dynamique dans le document.

Les parties XML personnalisées peuvent être associées à des espaces de noms. Pour obtenir des données à partir des parties XML personnalisées dans un espace de noms, utilisez la méthode [CustomXmlParts.getByNamespaceAsync](../../reference/shared/customxmlparts.getbynamespaceasync.md).

Vous pouvez également utiliser la [CustomXmlParts.getByIdAsync](../../reference/shared/customxmlparts.getbyidasync.md) pour accéder aux parties XML personnalisées par leur GUID. Après avoir obtenu une partie XML personnalisée, utilisez la méthode [CustomXmlPart.getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md) pour obtenir les données XML.

Pour ajouter une partie XML personnalisée à un document, utilisez la propriété **Document.customXmlParts** afin d’obtenir les parties XML personnalisées qui sont dans le document et appelez la méthode [CustomXmlParts.addAsync](../../reference/shared/customxmlparts.addasync.md).

Pour obtenir des informations détaillées sur l’utilisation de parties XML personnalisées avec un complément du volet Office, voir [Création de meilleurs compléments pour Word avec Office Open XML](../../docs/word/create-better-add-ins-for-word-with-office-open-xml.md).


## <a name="persisting-add-in-settings"></a>Persistance des paramètres de complément


Vous devez souvent enregistrer les données personnalisées pour votre complément, telles que les préférences d’un utilisateur ou l’état du complément, et accéder à ces données lors de la prochaine ouverture du complément. Vous pouvez utiliser des techniques de programmation web courantes pour enregistrer les données, comme les cookies de navigateur ou le stockage web HTML 5. Si votre complément est également exécuté dans Excel, PowerPoint ou Word, vous pouvez également utiliser les méthodes de l’objet [Settings](../../reference/shared/settings.md). Les données créées avec l’objet **Settings** sont stockées dans la feuille de calcul, la présentation ou le document dans lequel le complément a été inséré et enregistré. Ces données sont disponibles seulement pour le complément qui les a créées.

Pour éviter les allers-retours vers le serveur sur lequel le document est stocké, les données créées avec l’objet **Settings** sont gérées dans la mémoire lors de l’exécution. Les données de paramètres enregistrées précédemment sont chargées en mémoire lors de l’initialisation du complément et les modifications apportées à ces données sont uniquement enregistrées dans le document quand vous appelez la méthode [Settings.saveAsync](../../reference/shared/settings.saveasync.md). En interne, les données sont stockées dans un objet JSON sérialisé en tant que paires nom/valeur. Vous pouvez utiliser les méthodes [get](../../reference/shared/settings.get.md), [set](../../reference/shared/settings.set.md) et [remove](../../reference/shared/settings.removehandlerasync.md) de l’objet **Settings** pour lire, écrire et supprimer des éléments dans la copie en mémoire des données. La ligne de code suivante explique comment créer un paramètre nommé `themeColor` et définir sa valeur sur « green ».




```js
Office.context.document.settings.set('themeColor', 'green');
```

Étant donné que les données de paramètres créées ou supprimées avec les méthodes **set** et **remove** agissent sur une copie en mémoire des données, vous devez appeler **saveAsync** pour rendre persistantes les modifications apportées aux données de paramètres dans le document utilisé par votre complément.

Pour plus de détails sur l’utilisation des données personnalisées à l’aide des méthodes de l’objet **Settings**, voir la page relative à la [conservation des données et aux paramètres d’état de complément](../../docs/develop/persisting-add-in-state-and-settings.md).


## <a name="reading-properties-of-a-project-document"></a>Lecture des propriétés d’un document de projet

Si votre complément de volet Office s’exécute dans Project, vous pouvez lire les données de certains champs, ressources et champs de tâche du projet actif. Pour ce faire, vous pouvez utiliser les méthodes et les événements de l’objet [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md), qui étend l’objet **Document** pour fournir des fonctionnalités supplémentaires propres à Project.

Pour des exemples de lecture de données Project, voir [Créer votre premier complément du volet Office pour Projet 2013 à l’aide d’un éditeur de texte](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## <a name="permissions-model-and-governance"></a>Modèle d’autorisations et gouvernance

Votre complément utilise l’élément **Permissions** dans son manifeste pour demander l’autorisation d’accéder au niveau de fonctionnalité qu’il exige à partir de l’interface API JavaScript pour Office. Par exemple, si votre complément nécessite un accès en lecture/écriture pour le document, son manifeste doit spécifier `ReadWriteDocument` en tant que valeur de texte dans l’élément **Permissions**. Étant donné que les autorisations ont pour objectif de protéger la vie privée et la sécurité de l’utilisateur, en tant que meilleures pratiques, nous vous recommandons de demander le niveau d’autorisation minimal requis pour ses fonctionnalités. L’exemple suivant illustre la demande de l’autorisation **ReadDocument** dans le manifeste d’un complément du volet Office.


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
 xsi:type="TaskPaneApp">
???<!-- Other manifest elements omitted. -->
  <Permissions>ReadDocument</Permissions>
???
</OfficeApp>

```

Pour plus d’informations, consultez la page relative à la [demande d’autorisations pour l’utilisation de l’API dans des compléments de contenu et de volet Office](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).


## <a name="additional-resources"></a>Ressources supplémentaires


- [API JavaScript pour Office](../../reference/javascript-api-for-office.md)
    
- [Informations de référence sur le schéma des manifestes des applications pour Office](http://msdn.microsoft.com/en-us/library/7e0cadc3-f613-8eb9-57ef-9032cbb97f92.aspx)
    
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../../docs/testing/testing-and-troubleshooting.md)
    
