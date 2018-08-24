---
title: Demande d’autorisations d’utilisation de l’API dans des compléments de contenu et de volet des tâches
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: c7f303b1df20fedb41400d9b1f44512a2c5be579
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925359"
---
# <a name="requesting-permissions-for-api-use-in-content-and-task-pane-add-ins"></a>Demande d’autorisations d’utilisation de l’API dans des compléments de contenu et de volet des tâches

Cet article décrit les différents niveaux d’autorisation que vous pouvez déclarer dans le manifeste de votre complément de contenu ou du volet Office afin de spécifier le niveau d’accès de l’API JavaScript requis pour les fonctionnalités de votre complément. 




## <a name="permissions-model"></a>Modèles d’autorisation


Le modèle d’autorisations d’accès de l’API JavaScript à cinq niveaux fournit les bases de confidentialité et de sécurité pour les utilisateurs de vos compléments de contenu et du volet Office. La figure 1 montre les cinq niveaux d’autorisations de l’API que vous pouvez déclarer dans le manifeste de votre complément.


*Figure 1. Modèle d’autorisations à cinq niveaux pour les compléments de contenu et du volet Office*

![Niveaux d’autorisations des applications de volet de tâches](../images/office15-app-sdk-task-pane-app-permission.png)



Ces autorisations spécifient le sous-ensemble de l’API auquel votre complément de contenu ou du volet Office est autorisé à accéder par l’environnement d’exécution lorsqu’un utilisateur insère, puis active (approuve) votre complément. Pour déclarer le niveau d’autorisation nécessaire à votre complément de contenu ou du volet Office, indiquez l’une des valeurs de texte d’autorisation dans l’élément [Permissions](https://dev.office.com/reference/add-ins/manifest/permissions) du manifeste de votre complément. L’exemple suivant demande l’autorisation **WriteDocument**, laquelle n’autorise que les méthodes pouvant écrire dans le document (et non le lire).




```XML
<Permissions>WriteDocument</Permissions>
```

Il est recommandé de toujours demander les autorisations sur la base du principe d’ _autorisation minimum_. En d’autres termes, vous devez demander l’autorisation d’accéder uniquement au sous-ensemble de l’API nécessaire au bon fonctionnement de votre complément. Par exemple, si votre complément est conçu pour uniquement lire des données dans le document d’un utilisateur, vous ne devez demander que l’autorisation **ReadDocument**.

Le tableau suivant décrit le sous-ensemble de l’API JavaScript activé pour chaque niveau d’autorisation.



|**Autorisation**|**Sous-ensemble de l’API activé**|
|:-----|:-----|
|**Limité**|Les méthodes de l’objet [Settings](https://dev.office.com/reference/add-ins/shared/settings) et la méthode [Document.getActiveViewAsync](https://dev.office.com/reference/add-ins/shared/document.getactiveviewasync).Niveau d’autorisation minimal pouvant être demandé par un complément de contenu ou du volet Office.|
|**ReadDocument**|En plus de l’API activé par l’autorisation  **Restricted**, permet l’accès aux membres de l’API nécessaires à la lecture du document et à la gestion des liaisons.Cela inclut l’utilisation des éléments suivants :<br/><ul><li>La méthode <a href="https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync" target="_blank">Document.getSelectedDataAsync</a> pour obtenir les données texte, HTML (Word uniquement) ou tabulaires sélectionnées, mais pas le code sous-jacent Open Office XML (OOXML) contenant toutes les données du document.</p></li><li><p>La méthode <a href="https://dev.office.com/reference/add-ins/shared/document.getfileasync" target="_blank">Document.getFileAsync</a> pour l’obtention de la totalité du texte du document, mais pas la copie binaire OOXML sous-jacente du document.</p></li><li><p>La méthode <a href="https://dev.office.com/reference/add-ins/shared/binding.getdataasync" target="_blank">Binding.getDataAsync</a> pour la lecture des données liées dans le document.</p></li><li><p>Les méthodes <a href="https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync" target="_blank">addFromNamedItemAsync</a>, <a href="https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync" target="_blank">addFromPromptAsync</a> et <a href="https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync" target="_blank">addFromSelectionAsync</a> de l’objet <span class="keyword">Bindings</span> pour la création de liaisons dans le document.</p></li><li><p>Les méthodes <a href="https://dev.office.com/reference/add-ins/shared/bindings.getallasync" target="_blank">getAllAsync</a>, <a href="https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync" target="_blank">getByIdAsync</a> et <a href="https://dev.office.com/reference/add-ins/shared/bindings.releasebyidasync" target="_blank">releaseByIdAsync</a> de l’objet <span class="keyword">Bindings</span> pour accéder aux liaisons du document et les supprimer.</p></li><li><p>La méthode <a href="https://dev.office.com/reference/add-ins/shared/document.getfilepropertiesasync" target="_blank">Document.getFilePropertiesAsync</a> pour accéder aux propriétés du fichier de document, comme l’URL du document.</p></li><li><p>La méthode <a href="https://dev.office.com/reference/add-ins/shared/document.gotobyidasync" target="_blank">Document.goToByIdAsync</a> pour accéder aux objets et aux emplacements nommés dans le document.</p></li><li><p>Pour les compléments du volet Office de Project, toutes les méthodes d’obtention (get) de l’objet <a href="https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument" target="_blank">ProjectDocument</a>. </p></li></ul>|
|**ReadAllDocument**|En plus de l’API activé par les autorisations **Restricted** et **ReadDocument**, permet l’accès supplémentaire aux données de document ci-dessous :<br/><ul><li><p>Les méthodes <span class="keyword">Document.getSelectedDataAsync</span> et <span class="keyword">Document.getFileAsync</span> pour accéder au code OOXML sous-jacent du document (qui peut inclure une mise en forme, des liens, des graphiques incorporés, des commentaires, des révisions, etc. en plus du texte).</p></li></ul>|
|**WriteDocument**|En plus de l’API activé par l’autorisation **Restricted**, permet l’accès aux membres de l’API suivants :<br/><ul><li><p>La méthode <a href="https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync" target="_blank">Document.setSelectedDataAsync</a> pour écrire dans la sélection de l’utilisateur dans le document.</p></li></ul>|
|**ReadWriteDocument**|En plus de l’API activée par les autorisations  **Restricted**,  **ReadDocument**,  **ReadAllDocument** et **WriteDocument**, permet l’accès à toutes les API restantes prises en charge par les compléments de contenu et du volet Office, y compris les méthodes d’abonnement à des événements.Vous devez déclarer l’autorisation  **ReadWriteDocument** pour accéder aux membres supplémentaires suivants de l’API :<br/><ul><li><p>La méthode <a href="https://dev.office.com/reference/add-ins/shared/binding.setdataasync" target="_blank">Binding.setDataAsync</a> pour écrire dans des zones liées du document.</p></li><li><p>La méthode <a href="https://dev.office.com/reference/add-ins/shared/binding.tablebinding.addrowsasync" target="_blank">TableBinding.addRowsAsync</a> pour ajouter des lignes dans les tables liées.</p></li><li><p>La méthode <a href="https://dev.office.com/reference/add-ins/shared/binding.tablebinding.addcolumnsasync" target="_blank">TableBinding.addColumnsAsync</a> pour ajouter des colonnes dans les tables liées.</p></li><li><p>La méthode <a href="https://dev.office.com/reference/add-ins/shared/binding.tablebinding.deletealldatavaluesasync" target="_blank">TableBinding.deleteAllDataValuesAsync</a> pour supprimer toutes les données d’une table liée.</p></li><li><p>Les méthodes <a href="https://dev.office.com/reference/add-ins/shared/binding.tablebinding.setformatsasync" target="_blank">setFormatsAsync</a>, <a href="https://dev.office.com/reference/add-ins/shared/binding.tablebinding.clearformatsasync" target="_blank">clearFormatsAsync</a> et <a href="https://dev.office.com/reference/add-ins/shared/binding.tablebinding.settableoptionsasync" target="_blank">setTableOptionsAsync</a> de l’objet <span class="keyword">TableBinding</span> pour définir la mise en forme et les options des tables liées.</p></li><li><p>Tous les membres des objets <a href="https://dev.office.com/reference/add-ins/shared/customxmlnode.customxmlnode" target="_blank">CustomXmlNode</a>, <a href="https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart" target="_blank">CustomXmlPart</a>, <a href="https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts" target="_blank">CustomXmlParts</a> et <a href="https://dev.office.com/reference/add-ins/shared/customxmlprefixmappings.customxmlprefixmappings" target="_blank">CustomXmlPrefixMappings</a>.</p></li><li><p>Toutes les méthodes d’abonnement à des événements prises en charge par les compléments de contenu et du volet Office, en particulier les méthodes <span class="keyword">addHandlerAsync</span> et <span class="keyword">removeHandlerAsync</span> des objets <a href="https://dev.office.com/reference/add-ins/shared/binding" target="_blank">Binding</a>, <a href="https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart" target="_blank">CustomXmlPart</a>, <a href="https://dev.office.com/reference/add-ins/shared/document" target="_blank">Document</a>, <a href="https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument" target="_blank">ProjectDocument</a> et <a href="https://dev.office.com/reference/add-ins/shared/document.settings" target="_blank">Settings</a>.</p></li></ul>|

## <a name="see-also"></a>Voir aussi

- [Confidentialité et sécurité pour les compléments Office](../concepts/privacy-and-security.md)
    


