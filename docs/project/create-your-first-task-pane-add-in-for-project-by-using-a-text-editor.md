---
title: Créer votre premier complément de volet des tâches pour Microsoft Project à l’aide d’un éditeur de texte
description: Créez un complément du volet Office pour Project Standard 2013, Project Professionnel 2013 ou versions ultérieures à l’aide du générateur Yeoman pour les compléments Office.
ms.date: 07/10/2020
ms.localizationpriority: medium
ms.openlocfilehash: 69353b94da05dd0a8cfd6347beb7c5b3f4fd81e0
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66959027"
---
# <a name="create-your-first-task-pane-add-in-for-microsoft-project-by-using-a-text-editor"></a>Créer votre premier complément de volet des tâches pour Microsoft Project à l’aide d’un éditeur de texte

Vous pouvez créer un complément du volet Office pour Project Standard 2013, Project Professionnel 2013 ou versions ultérieures à l’aide du générateur Yeoman pour les compléments Office. Cet article explique comment créer un complément simple qui utilise un manifeste XML qui pointe vers un fichier HTML sur un partage de fichiers. L’exemple de complément Project OM Test teste certaines fonctions JavaScript qui utilisent le modèle objet pour les compléments. Après avoir utilisé le Centre de gestion de la **confidentialité** dans Project pour inscrire le partage de fichiers qui contient le fichier manifeste, vous pouvez ouvrir le complément du volet Office à partir de l’onglet **Projet** du ruban. (Le code de cet article est basé sur un exemple développé par Arvind Iyer, Microsoft Corporation.)

Project utilise le même schéma de manifeste de complément que celui utilisé par d’autres clients Office et une grande partie de la même API JavaScript. Le code complet pour le complément décrit dans cet article est disponible dans le `Samples\Apps` sous-répertoire du téléchargement SDK Project 2013.

L’exemple de complément Test du modèle objet de Project peut obtenir le GUID d’une tâche et les propriétés de l’application et du projet actif. Si Project Professionnel 2013 ouvre un projet se trouvant dans une bibliothèque SharePoint, le complément peut afficher l’URL du projet.

Le [téléchargement du kit de développement logiciel Project 2013](https://www.microsoft.com/download/details.aspx?id=30435) inclut le code source complet. Lorsque vous extrayez et installez le kit de développement logiciel et les exemples figurant dans le fichier Project2013SDK.msi, reportez-vous au sous-répertoire `\Samples\Apps\Copy_to_AppManifests_FileShare` pour le fichier manifeste et au sous-répertoire `\Samples\Apps\Copy_to_AppSource_FileShare` pour le code source.

L’exemple JSOMCall.html utilise les fonctions JavaScript du fichier office.js et le fichier project-15.js qui sont inclus. Vous pouvez utiliser les fichiers de débogage correspondants (office.debug.js et project-15.debug.js) pour examiner les fonctions.

Pour une introduction à l’utilisation de JavaScript dans les compléments Office, consultez [Présentation de l’API JavaScript Office](../develop/understanding-the-javascript-api-for-office.md).

## <a name="procedure-1-to-create-the-add-in-manifest-file"></a>Procédure 1. Pour créer le fichier de manifeste du complément

Créez un fichier XML dans un répertoire local. Le fichier XML inclut l’élément `OfficeApp` et les éléments enfants, qui sont [décrits dans le manifeste XML des compléments Office](../develop/add-in-manifests.md). Par exemple, créez un fichier nommé JSOM_SimpleOMCalls.xml qui contient le code XML suivant (modifiez la valeur GUID de l’élément `Id` ).

```XML
<?xml version="1.0" encoding="utf-8"?>
   <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
              xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
              xsi:type="TaskPaneApp">
     <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
     <Id>93A26520-9414-492F-994B-4983A1C7A607</Id>
     <Version>15.0</Version>
     <ProviderName>Microsoft</ProviderName>
     <DefaultLocale>en-us</DefaultLocale>
     <DisplayName DefaultValue="Project OM Test">
       <Override Locale="fr-fr" Value="Le Project OM Test"/>
     </DisplayName>
     <Description DefaultValue="Test the task pane add-in object model for Project - English (US)">
       <Override Locale="fr-fr" Value="Test the task pane add-in object model for Project - French (France)"/>
     </Description>
     <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
     <Hosts>
       <Host Name="Project"/>
       <Host Name="Workbook"/>
       <Host Name="Document"/>
     </Hosts>
    <DefaultSettings>
       <SourceLocation DefaultValue="\\ServerName\AppSource\JSOMCall.html">
         <Override Locale="fr-fr" Value="\\ServerName\AppSource\JSOMCall.html"/>
       </SourceLocation>
     </DefaultSettings>
     <Permissions>ReadWriteDocument</Permissions>
     <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg">
       <Override Locale="fr-fr" Value="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg"/>
     </IconUrl>
     <AllowSnapshot>true</AllowSnapshot>
   </OfficeApp>
```

Pour Project, l’élément `OfficeApp` doit inclure la valeur d’attribut `xsi:type="TaskPaneApp"` . L’élément `Id` est un GUID. La `SourceLocation` valeur doit être un chemin d’accès de partage de fichiers ou une URL SharePoint pour le fichier source HTML du complément ou l’application web qui s’exécute dans le volet Office. For an explanation of the other elements in manifest file, see [Task pane add-ins for Project](../project/project-add-ins.md).

La procédure 2 montre comment créer le fichier HTML que le manifeste JSOM_SimpleOMCalls.xml spécifie pour le complément de test de Project. Les boutons qui sont spécifiés dans le fichier HTML appellent des fonctions JavaScript associées. Vous pouvez ajouter les fonctions JavaScript dans le fichier HTML ou les placer dans un fichier .js distinct.

## <a name="procedure-2-to-create-the-source-files-for-the-project-om-test-add-in"></a>Procédure 2. Pour créer les fichiers sources du complément Test du modèle objet de Project

1. Créez un fichier HTML avec un nom spécifié par l’élément `SourceLocation` dans le manifeste JSOM_SimpleOMCalls.xml.

   Par exemple, créez le fichierJSOMCall.html dans le répertoire `C:\Project\AppSource`. Bien que vous puissiez utiliser un éditeur de texte simple pour créer les fichiers sources, il est plus facile d’utiliser un outil tel que Visual Studio Code, qui fonctionne avec des types de documents spécifiques (tels que HTML et JavaScript) et dispose d’autres aides à la modification. Si vous n’avez pas déjà effectué l’exemple de recherche Bing décrit dans [Compléments volet des tâches pour Project](../project/project-add-ins.md), la procédure 3 explique comment créer le `\\ServerName\AppSource` partage de fichiers spécifié par le manifeste.

   Le fichier JSOMCall.html utilise le fichier MicrosoftAjax.js commun pour la fonctionnalité AJAX et le fichier Office.js pour la fonctionnalité de complément dans les applications Office 2013.

    ```HTML
    <!DOCTYPE html>
    <html>
        <head>
            <title>Project OM Sample Code</title>
            <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
            <script type="text/javascript" src="MicrosoftAjax.js"></script>

            <!-- Use the CDN reference to office.js when deploying your add-in. -->
            <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script> -->
            <script type="text/javascript" src="office.js"></script>
            <script type="text/javascript" src="JSOM_Sample.js"></script>
        </head>
        <body>
            <div id="Common_JSOM_API">
                OBJECT MODEL TESTS
            </div>

            <textarea id="text" rows="6" cols="25">This is the text result.</textarea>
        </body>
    </html>
    ```

   L’élément `textarea` spécifie une zone de texte qui affiche les résultats des fonctions JavaScript.

   > [!NOTE]
   > Pour garantir le bon fonctionnement de l’exemple Test du modèle objet de Project, copiez les fichiers suivants du téléchargement du kit de développement logiciel de Project 2013 dans le même répertoire que le fichier JSOMCall.html : Office.js, Project-15.js et MicrosoftAjax.js.

   L’étape 2 ajoute le fichier JSOM_Sample.js pour les fonctions spécifiques que l’exemple de complément Test du modèle objet de Project utilise. Dans les dernières étapes, vous ajoutez d’autres éléments HTML pour les boutons qui appellent les fonctions JavaScript.

1. Créez un fichier JavaScript nommé JSOM_Sample.js dans le même répertoire que le fichier JSOMCall.html.

   Le code suivant obtient le contexte d’application et les informations de document à l’aide de méthodes dans le fichier Office.js. L’objet `text` est l’ID du `textarea` contrôle dans le fichier HTML.

   La **\_variable projDoc** est initialisée avec un `ProjectDocument` objet. Le code inclut certaines fonctions de gestion des erreurs simples, ainsi que la fonction qui obtient le `getContextValues` contexte d’application et les propriétés de contexte du document de projet. Pour plus d’informations sur le modèle d’objet JavaScript pour Project, voir [API JavaScript pour Office](../reference/javascript-api-for-office.md).

    ```js
    /*
    * JavaScript functions for the Project OM Test example app
    * in the Project 2013 SDK.
    */

    var _projDoc;
    var _app;
    var taskGuid;
    var resourceGuid;

    // The initialize function is required for all add-ins.
    Office.initialize = function (reason) {
        // Checks for the DOM to load using the jQuery ready method.
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
            _projDoc = Office.context.document;
            _app = Office.context;
        });
    }

    function logError(errorText) {
        text.value = "Error in " + errorText;
    }

    function logEventError(erroneousEvent) {
        logError("event " + erroneousEvent);
    }

    function logMethodError(methodName, errorName, errorMessage) {
        logError(methodName + " method.\nError name: " + errorName + "\nMessage: " + errorMessage);
    }

    // . . . Add other JavaScript functions here.

    function getContextValues() {
        getDocumentUrl();
        getDocumentMode();
        getApplicationContentLanguage();
        getApplicationDisplayLanguage();
    }

    function getDocumentUrl() {
        text.value ="Document URL:\n" +_projDoc.url;
    }

    function getDocumentMode() {
        var docMode = _projDoc.mode;
        text.value = text.value + "\n\nDocument mode: " + docMode;
    }

    function getApplicationContentLanguage() {
        text.value = text.value + "\nApp language: " + _app.contentLanguage;
    }

    function getApplicationDisplayLanguage() {
        text.value = text.value + "\nDisplay language: " + _app.displayLanguage;
    }
    ```

   Pour plus d’informations sur les fonctions dans le fichier Office.debug.js, consultez [l’API JavaScript Office](../reference/javascript-api-for-office.md). Par exemple, la `getDocumentUrl` fonction obtient l’URL ou le chemin d’accès au fichier du projet ouvert.

1. Ajoutez des fonctions JavaScript qui appellent des méthodes asynchrones dans Office.js et Project-15.js pour obtenir les données sélectionnées :

   - Par exemple, `getSelectedDataAsync` est une méthode générale dans Office.js qui obtient du texte non mis en forme pour les données sélectionnées. Pour plus d’informations, consultez [l’objet Document](/javascript/api/office/office.document#office-office-document-getselectedtaskasync-member(1)).

   - La `getSelectedTaskAsync` fonction dans Project-15.js obtient le GUID de la tâche sélectionnée. De même, la `getSelectedResourceAsync` fonction obtient le GUID de la ressource sélectionnée. Si vous appelez ces fonctions lorsqu’une tâche ou une ressource n’est pas sélectionnée, les fonctions produisent une erreur non définie.

   - La `getTaskAsync` fonction obtient le nom de la tâche et les noms des ressources affectées. Si la tâche se trouve dans une liste de tâches SharePoint synchronisée, `getTaskAsync` obtient l’ID de tâche dans la liste SharePoint ; sinon, l’ID de tâche SharePoint est 0.

     > [!NOTE]
     > À des fins de démonstration, l’exemple de code comporte un bogue. S’il `taskGuid` n’est pas défini, la `getTaskAsync` fonction se déconnecte. Si vous obtenez un GUID de tâche valide, puis sélectionnez une autre tâche, la `getTaskAsync` fonction obtient des données pour la tâche la plus récente qui a été opérée par la `getSelectedTaskAsync` fonction.
  
   - `getTaskFields`, `getResourceFields`et sont des `getProjectFields` fonctions locales qui appellent `getTaskFieldAsync`, `getResourceFieldAsync`ou `getProjectFieldAsync` plusieurs fois pour obtenir les champs spécifiés d’une tâche ou d’une ressource. Dans le fichier project-15.debug.js, l’énumération `ProjectTaskFields` et l’énumération `ProjectResourceFields` indiquent les champs pris en charge.

   - La `getSelectedViewAsync` fonction obtient le type de vue (défini dans l’énumération `ProjectViewTypes` dans project-15.debug.js) et le nom de la vue.

   - Si le projet est synchronisé avec une liste de tâches SharePoint, la `getWSSUrlAsync` fonction obtient l’URL et le nom de la liste de tâches. Si le projet n’est pas synchronisé avec une liste de tâches SharePoint, la `getWSSUrlAsync` fonction se termine par des erreurs.

     > [!NOTE]
     > Pour obtenir l’URL SharePoint et le nom de la liste des tâches, nous vous recommandons d’utiliser la `getProjectFieldAsync` méthode avec les constantes et `WSSList` les `WSSUrl` constantes dans l’énumération [ProjectProjectFields](/javascript/api/office/office.projectprojectfields).

   Chacune des fonctions utilisées dans le code suivant inclut une fonction anonyme représentée par `function (asyncResult)` et qui est un rappel qui obtient le résultat asynchrone. Au lieu de fonctions anonymes, vous pouvez utiliser les fonctions nommées, qui peuvent améliorer la maintenabilité des compléments complexes.

    ```js
    // Get the data in the selected cells of the grid in the active view.
    function getSelectedDataAsync() {
        _projDoc.getSelectedDataAsync(
            Office.CoercionType.Text,
            { ValueFormat: "Formatted" },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded)
                    text.value = asyncResult.value;
                else
                    logMethodError("getSelectedDataAsync", asyncResult.error.name,
                                   asyncResult.error.message);
            }
        );
    }

    // Get the GUID of the selected task.
    function getSelectedTaskAsync() {
        _projDoc.getSelectedTaskAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = asyncResult.value;
                taskGuid = asyncResult.value;
            }
            else {
                logMethodError("getSelectedTaskAsync", asyncResult.error.name,
                                   asyncResult.error.message);
            }
        });
    }

    // Get the GUID of the selected resource.
    function getSelectedResourceAsync() {
        _projDoc.getSelectedResourceAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = asyncResult.value;
                resourceGuid = asyncResult.value;
            }
            else {
                logMethodError("getSelectedResourceAsync", asyncResult.error.name,
                                   asyncResult.error.message);
            }
        });
    }

    // Get data for the specified task.
    function getTaskAsync() {
        if (taskGuid != undefined) {
            _projDoc.getTaskAsync(
                taskGuid,
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        logMethodError("getTaskAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                    } else {
                        var taskInfo = asyncResult.value;
                        var taskOutput = "Task name: " + taskInfo.taskName +
                                         "\nGUID: " + taskGuid +
                                         "\nWSS Id: " + taskInfo.wssTaskId +
                                         "\nResourceNames: " + taskInfo.resourceNames;
                        text.value = taskOutput;
                    }
                }
            );
        } else {
            text.value = 'Task GUID not valid:\n' + taskGuid;
        } 
    }

    // Get additional data for task fields.
    function getTaskFields() {
        text.value = "";

        _projDoc. getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Name,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Name: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.ID,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "ID: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Start,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Start: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Duration,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Duration: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Priority,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Priority: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Notes,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Notes: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        ); 
    }

    // Get data for the specified resource fields.
    function getResourceFields() {
        text.value = "";

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.Name,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Resource name: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.Cost,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Cost: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.StandardRate,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Standard Rate: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.ActualCost,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Actual Cost: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.ActualWork,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Actual Work: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.Units,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Units: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );
    }

    // Get the URL and list name of the synchronized SharePoint task list.
    // Recommended: use getProjectField instead.
    function getWSSUrlAsync() {
        _projDoc.getWSSUrlAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = "SharePoint URL:\n" + asyncResult.value.serverUrl
                    + "\nList name: " + asyncResult.value.listName;
            }
            else {
                logMethodError("getWSSUrlAsync", asyncResult.error.name, asyncResult.error.message);
            }
        });
    }

    // Get the type and name of the selected view.
    function getSelectedViewAsync() {
        _projDoc.getSelectedViewAsync(function (asyncResult) {
            text.value = "View type: " + asyncResult.value.viewType
                + "\nName: " + asyncResult.value.viewName;
        });
    }

    // Get information about the active project.
    function getProjectFields() {
        text.value = "";

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.GUID,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Project GUID: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.Start,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nStart: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.Finish,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nFinish: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProject " + errorText);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.CurrencyDigits,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nCurrency digits: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.CurrencySymbol,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Currency symbol: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.CurrencySymbolPosition,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nSymbol position: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.ProjectServerUrl,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nProject web app URL:\n  " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.WSSUrl,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nSharePoint URL:\n  " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.WSSList,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nSharePoint list: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }
    ```

1. Ajoutez des rappels et des fonctions du gestionnaire d’événements JavaScript pour enregistrer la sélection de tâches, la sélection de ressources et les gestionnaires d’événements de changement de sélection d’affichage, et pour annuler l’enregistrement de gestionnaires d’événements. La `manageEventHandlerAsync` fonction ajoute ou supprime le gestionnaire d’événements spécifié, en fonction du paramètre *d’opération* . L’opération peut être `addHandlerAsync` ou `removeHandlerAsync`.

   Les `manageTaskEventHandler`fonctions , `manageResourceEventHandler`et `manageViewEventHandler` peuvent ajouter ou supprimer un gestionnaire d’événements, comme spécifié par le paramètre *docMethod* .

    ```js
    // Task selection changed event handler.
    function onTaskSelectionChanged(eventArgs) {
        text.value = "In task selection change event handler";
    }

    // Resource selection changed event handler.
    function onResourceSelectionChanged(eventArgs) {
        text.value = "In Resource selection changed event handler";
    }

    // View selection changed event handler.
    function onViewSelectionChanged(eventArgs) {
        text.value = "In View selection changed event handler";
    }

    // Add or remove the specified event handler.
    function manageEventHandlerAsync(eventType, handler, operation, onComplete) {
        _projDoc[operation]   //The operation is addHandlerAsync or removeHandlerAsync.
        (
            eventType,
            handler,
            function (asyncResult) {
                if (onComplete) {
                    onComplete(asyncResult, operation);
                } else {
                    var message = "Operation: " + operation;
                    message = message + "\nStatus: " + asyncResult.status + "\n";
                    text.value = message;
                }
            }
        );
    }

    // Write the asyncResult status from the manageEventHandlerAsync function (optional). 
    function onComplete(asyncResult, operation) {
        var message = "In onComplete function for " + operation;
        message = message + "\nStatus: " + asyncResult.status;
        text.value = message;
    }

    // Add or remove a task selection changed event handler.
    function manageTaskEventHandler(docMethod) {
        manageEventHandlerAsync(
            Office.EventType.TaskSelectionChanged,      // The task selection changed event.
            onTaskSelectionChanged,                     // The event handler.
            docMethod,                // The Office.Document method to add or remove an event handler.
            onComplete                // Manages the successful asyncResult data (optional).
        );
    }

    // Add or remove a resource selection changed event handler.
    function manageResourceEventHandler(docMethod) {
        manageEventHandlerAsync(
            Office.EventType.ResourceSelectionChanged,  // The resource selection changed event.
            onResourceSelectionChanged,                 // The event handler.
            docMethod,                // The Office.Document method to add or remove an event handler.
            onComplete                // Manages the successful asyncResult data (optional).
        );
    }

    // Add or remove a view selection changed event handler.
    function manageViewEventHandler(docMethod) {
        manageEventHandlerAsync(
            Office.EventType.ViewSelectionChanged,      // The view selection changed event.
            onViewSelectionChanged,                     // The event handler.
            docMethod,                // The Office.Document method to add or remove an event handler.
            onComplete                // Manages the successful asyncResult data (optional).
        );
    }
    ```

1. Pour le corps du document HTML, ajoutez des boutons qui appellent les fonctions JavaScript pour le test. Par exemple, dans l’élément `div` de l’API JSOM commune, ajoutez un bouton d’entrée qui appelle la fonction générale `getSelectedDataAsync` .

    ```HTML
    <body>
        <div id="Common_JSOM_API">
        OBJECT MODEL TESTS
        <br /><br />
        <strong>General function:</strong>
        <br />
        <input id="Button5" class="button-wide" type="button" onclick="getSelectedDataAsync()" 
            value="getSelectedDataAsync" />
        </div>
        <!--  more code . . .  -->
    ```

1. Ajoutez une `div` section avec des boutons pour les fonctions de tâche spécifiques au projet et pour l’événement `TaskSelectionChanged` .

    ```HTML
    <div id="ProjectSpecificTask">
      <br />
      <strong>Project-specific task methods:</strong><br />
      <button class="button-wide" onclick="getSelectedTaskAsync()">getSelectedTaskAsync</button><br />
      <button class="button-wide" onclick="getTaskAsync()">getTaskAsync</button><br />
      <button class="button-wide" onclick="getTaskFields()">Get Task Fields</button><br />
      <button class="button-wide" onclick="getWSSUrlAsync()">getWSSUrlAsync</button>
      <strong>Task selection changed:</strong>
      <button class="button-narrow" onclick="manageTaskEventHandler('addHandlerAsync')">Add</button>
      <button class="button-narrow" onclick="manageTaskEventHandler('removeHandlerAsync')">Remove</button>
    </div>
    ```

1. Ajouter `div` des sections avec des boutons pour les méthodes de ressources et les événements, afficher les méthodes et les événements, les propriétés du projet et les propriétés de contexte

    ```HTML
    <div id="ResourceMethods">
      <br />
      <strong>Resource methods:</strong>
      <button class="button-wide" onclick="getSelectedResourceAsync()">getSelectedResourceAsync</button><br />
      <button class="button-wide" onclick="getResourceFields()">Get Resource Fields</button><br />
      <strong>Resource selection changed:</strong>
      <button class="button-narrow" onclick="manageResourceEventHandler('addHandlerAsync')">Add</button>
      <button class="button-narrow" onclick="manageResourceEventHandler('removeHandlerAsync')">Remove</button>
    </div>
    <div id="ViewMethods">
      <br />
      <strong>View method:</strong>
      <button class="button-wide" onclick="getSelectedViewAsync()">getSelectedViewAsync</button><br />
      <strong>View selection changed:</strong>
      <button class="button-narrow" onclick="manageViewEventHandler('addHandlerAsync')">Add</button>
      <button class="button-narrow" onclick="manageViewEventHandler('removeHandlerAsync')">Remove</button>
    </div>
    <div id="ProjectMethods">
      <br />
      <strong>Project properties:</strong>
      <button class="button-wide" onclick="getProjectFields()">Get Project Fields</button><br />
    </div>
    <div id="ContextVariables">
      <br />
      <strong>Context properties:</strong>
      <button class="button-wide" onclick="getContextValues()">Get Context Values</button>
    </div>
    ```

1. Pour mettre en forme les éléments de bouton, ajoutez un élément CSS `style` . Par exemple, ajoutez ce qui suit en tant qu’enfant de l’élément `head` .

    ```HTML
    <style type="text/css">
        .button-wide
        {
            width: 210px;
            margin-top: 2px;
        }
        .button-narrow
        {
            width: 80px;
            margin-top: 2px;
        }
    </style>
    ```

La procédure 3 montre comment installer et utiliser les fonctionnalités du complément Test du modèle objet de Project.

## <a name="procedure-3-to-install-and-use-the-project-om-test-add-in"></a>Procédure 3. Pour installer et utiliser le complément Test du modèle objet de Project

1. Créez un partage de fichiers pour le répertoire qui contient le manifeste JSOM_SimpleOMCalls.xml. Vous pouvez créer le partage de fichiers sur l’ordinateur local ou sur un ordinateur distant accessible sur le réseau. Par exemple, si le manifeste se trouve dans le  `C:\Project\AppManifests` répertoire de l’ordinateur local, exécutez la commande suivante.

    `Net share AppManifests=C:\Project\AppManifests`

1. Créez un partage de fichiers pour le répertoire contenant les fichiers HTML et JavaScript pour le complément Test du modèle objet de Project. Assurez-vous que le chemin du partage de fichiers correspond à celui qui est spécifié dans le manifeste JSOM_SimpleOMCalls.xml. Par exemple, si les fichiers se trouvent dans le  `C:\Project\AppSource` répertoire de l’ordinateur local, exécutez la commande suivante.

    `net share AppSource=C:\Project\AppSource`

1. Dans Project, ouvrez la boîte de dialogue **Options de Project**, choisissez **Centre de gestion de la confidentialité**, puis choisissez **Paramètres du Centre de gestion de la confidentialité**.

   La procédure d’inscription d’un complément est également décrite dans la rubrique relative aux [compléments de volet Office pour Project](../project/project-add-ins.md), qui contient aussi des informations supplémentaires.

1. Dans la boîte de dialogue **Centre de gestion de la confidentialité**, dans le volet gauche, choisissez **Catalogues de compléments approuvés**.

1. Si vous avez déjà ajouté le `\\ServerName\AppManifests` chemin du complément Recherche Bing, ignorez cette étape. Sinon, dans le volet **Catalogues de compléments approuvés** , ajoutez le `\\ServerName\AppManifests` chemin d’accès dans la zone de texte **Url** du catalogue, choisissez **Ajouter un catalogue, activez** le partage réseau en tant que source par défaut (voir figure 1), puis choisissez **OK**.

   *Figure 1. Ajout d’un partage de fichiers réseau pour des manifestes de complément*

   ![Ajout d’un partage de fichiers réseau pour les manifestes d’application.](../images/pj15-create-simple-agave-manage-catalogs.png)

1. Après que vous avez ajouté de nouveaux compléments ou modifié le code source, redémarrez Project. Dans le ruban **PROJECT**, choisissez le menu déroulant **Compléments Office**, puis choisissez **Afficher tout**. Dans la boîte de dialogue **Insérer un complément**, choisissez **DOSSIER PARTAGÉ** (voir la figure 2), sélectionnez **Test du modèle objet de Project**, puis choisissez **Insérer**. Le complément Test du modèle objet de Project démarre dans un volet Office.

   *Figure 2. Démarrage du complément Test du modèle objet Project qui se trouve sur un partage de fichiers*

   ![Insertion d’une application.](../images/pj15-create-simple-agave-start-agave-app.png)

1. Dans Project, créez et enregistrez un projet simple comportant au moins deux tâches. Par exemple, créez les tâches nommées T1, T2 et un jalon nomméM1, puis définissez des durées et des prédécesseurs de tâches similaires à ceux de la figure 3. Choisissez l’onglet **PROJECT** sur le ruban, sélectionnez toute la ligne pour la tâche T2, puis cliquez sur le bouton **getSelectedDataAsync** dans le volet Office. La figure 3 montre les données qui sont sélectionnées dans la zone de texte du complément **Test du modèle objet de Project**.

   *Figure 3. Utilisation du complément Test du modèle objet Project*

   ![Utilisation de l’application Project OM Test.](../images/pj15-create-simple-agave-project-om-test.png)

1. Sélectionnez la cellule dans la colonne **Durée** de la première tâche, puis cliquez sur le bouton **getSelectedDataAsync** dans le complément **Test du modèle objet de Project**. La `getSelectedDataAsync` fonction définit la valeur de la zone de texte à afficher `2 days`.

1. Sélectionnez les trois cellules **Durée** pour les trois tâches. La `getSelectedDataAsync` fonction retourne des valeurs de texte séparées par des points-virgules pour les cellules sélectionnées dans différentes lignes, par exemple. `2 days;4 days;0 days`

   La `getSelectedDataAsync` fonction retourne des valeurs de texte séparées par des virgules pour les cellules sélectionnées dans une ligne. Par exemple, dans la figure 3, la ligne entière correspondant à la tâche T2 est sélectionnée. Lorsque vous choisissez `getSelectedDataAsync`, la zone de texte affiche les éléments suivants :  `,Auto Scheduled,T2,4 days,Thu 6/14/12,Tue 6/19/12,1,,<NA>`

   La colonne **Indicateurs** et la colonne **Noms des ressources** sont toutes deux vides, le tableau de texte affiche donc des valeurs vides pour ces colonnes. La valeur `<NA>` correspond à la cellule **Ajouter une nouvelle colonne**.

1. Sélectionnez une cellule dans la ligne de la tâche T2, ou toute la ligne pour la tâche T2, puis choisissez **getSelectedTaskAsync**. La zone de texte affiche la valeur GUID de la tâche, par exemple  `{25D3E03B-9A7D-E111-92FC-00155D3BA208}`. Project stocke cette valeur dans la variable globale `taskGuid` du complément **Project OM Test** .

1. Sélectionnez `getTaskAsync`. Si la `taskGuid` variable contient le GUID de la tâche T2, la zone de texte affiche les informations de tâche. La valeur **ResourceNames** est vide.

    Créez deux ressources locales R1 et R2, affectez-les à la tâche T2 à 50 % chacune, puis choisissez **à nouveau getTaskAsync** . Les résultats qui apparaissent dans la zone de texte incluent des informations sur les ressources. Si la tâche se trouve dans une liste de tâches SharePoint synchronisée, les résultats incluent également l’ID de tâche SharePoint.

    - Nom de la tâche : `T2`
    - GUID : `{25D3E03B-9A7D-E111-92FC-00155D3BA208}`
    - Identifiant de WSS : `0`
    - ResourceNames: `R1[50%],R2[50%]`

1. Sélectionnez le bouton **Obtenir les champs de tâche** . La `getTaskFields` fonction appelle la `getTaskFieldAsync` méthode plusieurs fois pour le nom de la tâche, l’index, la date de début, la durée, la priorité et les notes de tâche.

    - Nom : `T2`
    - Identifiant : `2`
    - Début : `Thu 6/14/12`
    - Durée : `4d`
    - Priorité : `500`
    - Remarques : il s’agit d’une note de la tâche T2. Il s’agit uniquement d’une note de test. Si elle avait été une note réelle, elle contiendrait des informations réelles.

1. Sélectionnez le bouton **getWSSUrlAsync**. Si le projet appartient à l’un des types suivants, les résultats présentent l’URL et le nom de la liste de tâches.

    - Une liste de tâches SharePoint qui a été importée dans Project Server.
    - Une liste de tâches SharePoint qui a été importée dans Project Professionnel, puis enregistrée à nouveau dans SharePoint (sans utiliser Project Server).

    > [!NOTE]
    > Si Project Professionnel est installé sur un ordinateur Windows Server, pour réenregistrer le projet dans SharePoint, vous pouvez utiliser le **Gestionnaire de serveur** afin d’ajouter la fonctionnalité **Expérience utilisateur**.

    Si le projet est un projet local ou si vous utilisez Project Professionnel pour ouvrir un projet géré par Project Server, la `getWSSUrlAsync` méthode affiche une erreur non définie.

    - URL SharePoint : `http://ServerName`
    - Nom de la liste : `Test task list`

1. Sélectionnez le bouton **Ajouter** dans la section **événement TaskSelectionChanged** , qui appelle la `manageTaskEventHandler` fonction pour inscrire un événement modifié de sélection de tâche et retourne `In onComplete function for addHandlerAsync Status: succeeded` dans la zone de texte. Sélectionnez une autre tâche ; la zone de texte affiche `In task selection changed event handler`, qui représente la sortie de la fonction de rappel pour l’événement de changement de sélection de tâche. Cliquez sur le bouton **Supprimer** pour annuler l’enregistrement du gestionnaire d’événements.

1. Pour utiliser des méthodes de ressources, sélectionnez d’abord un affichage tel que **Tableau des ressources**, **Utilisation des ressources** ou **Formulaire ressource**, puis sélectionnez une ressource dans cet affichage. Choisissez **getSelectedResourceAsync** pour initialiser la variable **resourceGuid** , puis **choisissez Obtenir les champs de ressources** à appeler `getResourceFieldAsync` plusieurs fois pour les propriétés de ressource. Vous pouvez également ajouter ou supprimer le gestionnaire d’événements de changement de sélection de ressources.

    - Nom de la ressource : `R1`
    - Coût : `$800.00`
    - Taux standard : `$50.00/h`
    - Coût réel : `$0.00`
    - Travail réel : `0h`
    - Unités : `100%`

1. Sélectionnez **getSelectedViewAsync** pour afficher le type et le nom de l’affichage actif. Vous pouvez également ajouter ou supprimer le gestionnaire d’événements de changement de sélection d’affichage. Par exemple, si **Resource Form** est l’affichage actif, la `getSelectedViewAsync` fonction affiche ce qui suit dans la zone de texte.

    - Type d’affichage : `6`
    - Nom : `Resource Form`

1. Sélectionnez **Obtenir les champs de projet** pour appeler la `getProjectFieldAsync` fonction plusieurs fois pour différentes propriétés du projet actif. Si le projet est ouvert à partir de Project Web App, la `getProjectFieldAsync` fonction peut obtenir l’URL de l’instance Project Web App.

    - GUID du projet : `9845922E-DAB4-E111-8AF3-00155D3BA208`
    - Début : `Tue 6/12/12`
    - Fin : `Tue 6/19/12`
    - Chiffres de la devise :`2`
    - Symbole monétaire : `$`
    - Position du symbole : `0`
    - URL Project Web App : `http://servername/pwa`
  
1. Sélectionnez le bouton **Obtenir les valeurs de contexte pour obtenir les propriétés** du document et de l’application dans laquelle le complément s’exécute, en obtenant les propriétés de l’objet **Office.Context.document** et de l’objet `Office.context.application` . Par exemple, si le fichier Project1.mpp se trouve sur le bureau de l’ordinateur local, l’URL du document est  `C:\Users\UserAlias\Desktop\Project1.mpp`. Si le fichier .mpp se trouve dans une bibliothèque SharePoint, la valeur est l’URL du document. Si vous utilisez Project Professionnel 2013 pour ouvrir un projet nommé Project1 à partir de Project Web App, l’URL du document est  `<>\Project1`.

    - URL du document : `<>\Project1`
    - Mode de document : `readWrite`
    - Langue de l’application : `en-US`
    - Langue d’affichage : `en-US`

1. Vous pouvez actualiser le complément après avoir édité le code source en fermant et en redémarrant Project. Dans le ruban **Project**, la liste déroulante **Compléments Office** contient la liste des compléments récemment utilisés.

## <a name="example"></a>Exemple

Le kit de développement logiciel Project 2013 contient le code complet du fichier JSOMCall.html, le fichier JSOM_Sample.js et les fichiers Office.js, Office.debug.js, Project-15.js et Project-15.debug.js associés. Voici le code du fichier JSOMCall.html.

```HTML
<!DOCTYPE html>
<html>
    <head>
        <title>Project OM Sample Code</title>
        <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>

        <script type="text/javascript" src="MicrosoftAjax.js"></script>

        <!-- Use the CDN reference to office.js when deploying your add-in. -->
        <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script> -->
        <script type="text/javascript" src="office.js"></script>
        <script type="text/javascript" src="JSOM_Sample.js"></script>

        <style type="text/css">
            .button-wide {
                width: 210px;
                margin-top: 2px;
            }
            .button-narrow 
            {
                width: 80px;
                margin-top: 2px;
            }
        </style>
    </head>

    <body>
        <div id="Common_JSOM_API">
            OBJECT MODEL TESTS
            <br /><br />
            <strong>General method:</strong>
            <br />
            <input id="Button5" class="button-wide" type="button" onclick="getSelectedDataAsync()" 
                value="getSelectedDataAsync" />
        </div>
        <div id="ProjectSpecificTask">
            <br />
            <strong>Project-specific task methods:</strong><br />
            <button class="button-wide" onclick="getSelectedTaskAsync()">getSelectedTaskAsync</button><br />
            <button class="button-wide" onclick="getTaskAsync()">getTaskAsync</button><br />
            <button class="button-wide" onclick="getTaskFields()">Get Task Fields</button><br />
            <button class="button-wide" onclick="getWSSUrlAsync()">getWSSUrlAsync</button>
            <strong>Task selection changed:</strong>
            <button class="button-narrow" onclick="manageTaskEventHandler('addHandlerAsync')">Add</button>
            <button class="button-narrow" onclick="manageTaskEventHandler('removeHandlerAsync')">Remove</button>
        </div>
        <div id="ResourceMethods">
            <br />
            <strong>Resource methods:</strong>
            <button class="button-wide" onclick="getSelectedResourceAsync()">getSelectedResourceAsync</button><br />
            <button class="button-wide" onclick="getResourceFields()">Get Resource Fields</button><br />
            <strong>Resource selection changed:</strong>
            <button class="button-narrow" onclick="manageResourceEventHandler('addHandlerAsync')">Add</button>
            <button class="button-narrow" onclick="manageResourceEventHandler('removeHandlerAsync')">Remove</button>
        </div>
        <div id="ViewMethods">
            <br />
            <strong>View method:</strong>
            <button class="button-wide" onclick="getSelectedViewAsync()">getSelectedViewAsync</button><br />
            <strong>View selection changed:</strong>
            <button class="button-narrow" onclick="manageViewEventHandler('addHandlerAsync')">Add</button>
            <button class="button-narrow" onclick="manageViewEventHandler('removeHandlerAsync')">Remove</button>
        </div>
        <div id="ProjectMethods">
            <br />
            <strong>Project properties:</strong>
            <button class="button-wide" onclick="getProjectFields()">Get Project Fields</button><br />
        </div>
        <div id="ContextVariables">
            <br />
            <strong>Context properties:</strong>
            <button class="button-wide" onclick="getContextValues()">Get Context Values</button>
        </div>
        <br />
        <textarea id="text" rows="10" cols="25">This is the text result.</textarea>
    </body>
</html
```

## <a name="robust-programming"></a>Programmation fiable

Le complément **Project OM Test** est un exemple qui montre l’utilisation de certaines fonctions JavaScript pour Project 2013 dans les fichiers Project-15.js et Office.js. L’exemple est destiné uniquement à des fins de test et n’inclut pas de contrôles d’erreur fiables. Par exemple, si vous ne sélectionnez pas de ressource et exécutez la `getSelectedResourceAsync` fonction, la `resourceGuid` variable n’est pas initialisée et appelle pour `getResourceFieldAsync` retourner une erreur. Pour un complément de production, vous devez vérifier l’absence d’erreurs spécifiques et ignorer les résultats, masquer la fonctionnalité qui ne s’applique pas ou avertir l’utilisateur de choisir une vue et d’effectuer une sélection valide avant d’utiliser une fonction.

Pour un exemple simple, la sortie d’erreur dans le code suivant inclut la th  `actionMessage` variable qui spécifie l’action à effectuer pour éviter une erreur dans la `getSelectedResourceAsync` fonction.

```js
function logError(errorText) {
    text.value = "Error in " + errorText;
}

function logMethodError(methodName, errorName, errorMessage, actionMessage) {
    logError(methodName + " method.\nError name: " + errorName
        + "\nMessage: " + errorMessage
        + "\n\nAction: " + actionMessage);
}

// Get the GUID of the selected resource.
function getSelectedResourceAsync() {
    _projDoc.getSelectedResourceAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
            text.value = asyncResult.value;
            resourceGuid = asyncResult.value;
        }
        else {
            var actionMessage = "Select a resource before running the getSelectedResourceAsync method.";
            logMethodError("getSelectedResourceAsync", asyncResult.error.name,
                               asyncResult.error.message, actionMessage);
        }
    });
}
```

L’exemple **HelloProject_OData** dans le téléchargement de Project 2013 SDK inclut le fichier SurfaceErrors.js qui utilise la bibliothèque JQuery pour afficher un message d’erreur contextuel. Figure 4 affiche le message d’erreur dans une notification d’annonce .

Le code suivant dans le fichier SurfaceErrors.js inclut la th  `throwError` function qui crée un `Toast` objet.

```js
/*
 * Show error messages in a "toast" notification.
 */

// Throws a custom defined error.
function throwError(errTitle, errMessage) {
    try {
        // Define and throw a custom error.
        var customError = { name: errTitle, message: errMessage }
        throw customError;
    }
    catch (err) {
        // Catch the error and display it to the user.
        Toast.showToast(err.name, err.message);
    }
}

// Add a dynamically-created div "toast" for displaying errors to the user.
var Toast = {

    Toast: "divToast",
    Close: "btnClose",
    Notice: "lblNotice",
    Output: "lblOutput",

    // Show the toast with the specified information.
    showToast: function (title, message) {

        if (document.getElementById(this.Toast) == null) {
            this.createToast();
        }

        document.getElementById(this.Notice).innerText = title;
        document.getElementById(this.Output).innerText = message;

        $("#" + this.Toast).hide();
        $("#" + this.Toast).show("slow");
    },

    // Create the display for the toast.
    createToast: function () {
        var divToast;
        var lblClose;
        var btnClose;
        var divOutput;
        var lblOutput;
        var lblNotice;

        // Create the container div.
        divToast = document.createElement("div");
        var toastStyle = "background-color:rgba(220, 220, 128, 0.80);" +
            "position:absolute;" +
            "bottom:0px;" +
            "width:90%;" +
            "text-align:center;" +
            "font-size:11pt;";
        divToast.setAttribute("style", toastStyle);
        divToast.setAttribute("id", this.Toast);

        // Create the close button.
        lblClose = document.createElement("div");
        lblClose.setAttribute("id", this.Close);
        var btnStyle = "text-align:right;" +
            "padding-right:10px;" +
            "font-size:10pt;" +
            "cursor:default";
        lblClose.setAttribute("style", btnStyle);
        lblClose.appendChild(document.createTextNode("CLOSE "));

        btnClose = document.createElement("span");
        btnClose.setAttribute("style", "cursor:pointer;");
        btnClose.setAttribute("onclick", "Toast.close()");
        btnClose.innerText = "X";
        lblClose.appendChild(btnClose);

        // Create the div to contain the toast title and message.
        divOutput = document.createElement("div");
        divOutput.setAttribute("id", "divOutput");
        var outputStyle = "margin-top:0px;";
        divOutput.setAttribute("style", outputStyle);

        lblNotice = document.createElement("span");
        lblNotice.setAttribute("id", this.Notice);
        var labelStyle = "font-weight:bold;margin-top:0px;";
        lblNotice.setAttribute("style", labelStyle);

        lblOutput = document.createElement("span");
        lblOutput.setAttribute("id", this.Output);

        // Add the child nodes to the toast div.
        divOutput.appendChild(lblNotice);
        divOutput.appendChild(document.createElement("br"));
        divOutput.appendChild(lblOutput);
        divToast.appendChild(lblClose);
        divToast.appendChild(divOutput);

        // Add the toast div to the document body.
        document.body.appendChild(divToast);
    },

    // Close the toast.
    close: function () {
        $("#" + this.Toast).hide("slow");
    }
}
```

Pour utiliser la `throwError` fonction, incluez la bibliothèque JQuery et le script SurfaceErrors.js dans le fichier JSOMCall.html, puis ajoutez un appel à `throwError` d’autres fonctions JavaScript telles que `logMethodError`.

> [!NOTE]
> Avant de déployer le complément, remplacez la référence à office.js et celle à jQuery par la référence au réseau de distribution de contenu. Cette dernière permet d’accéder à la version la plus récente et d’obtenir de meilleures performances.

```HTML
<!DOCTYPE html>
<html>
<head>
    <title>Project OM Sample Code</title>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />

    <script type="text/javascript" src="MicrosoftAjax.js"></script>

    <!-- Use the CDN reference to Office.js and jQuery when deploying your add-in. -->
    <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script> -->
    <script type="text/javascript" src="office.js"></script>
    <script type="text/javascript" src="http://ajax.microsoft.com/ajax/jQuery/jquery-1.9.0.min.js"></script>

    <script type="text/javascript" src="JSOM_Sample.js"></script>
    <script type="text/javascript" src="SurfaceErrors.js"></script>

    <!-- . . . INVALID USE OF SYMBOLS . . . -->
</head>
```

```js
function logMethodError(methodName, errorName, errorMessage, actionMessage) {
    logError(methodName + " method.\nError name: " + errorName
        + "\nMessage: " + errorMessage
        + "\n\nAction: " + actionMessage);

    throwError(methodName + " error", actionMessage);
}
```

*Figure 4. Les fonctions incluses dans le fichier SurfaceErrors.js peuvent afficher une notification « toast »*

![Utilisation des routines SurfaceError pour afficher une erreur.](../images/pj15-create-simple-agave-surface-error.png)

## <a name="see-also"></a>Voir aussi

- [Compléments du volet Office pour Project](../project/project-add-ins.md)
- [Présentation de l’API JavaScript pour compléments](../develop/understanding-the-javascript-api-for-office.md)
- [Compléments d’API JavaScript Office](../reference/javascript-api-for-office.md)
- [Informations de référence sur le schéma des manifestes des applications pour Office (version 1.1)](../develop/add-in-manifests.md)
- [Téléchargement du Kit de développement logiciel (SDK) de Project 2013](https://www.microsoft.com/download/details.aspx?id=30435)
