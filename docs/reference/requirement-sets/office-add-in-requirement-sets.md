# <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

L'ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou une vérification à l’exécution pour déterminer si un hôte Office prend en charge les API nécessaires à un complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Vous avez besoin d’informations sur l’endroit où les compléments sont pris en charge par l’hôte Office ? Consultez la page relative à la [Disponibilité des compléments Office sur les plateformes et les hôtes](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability).

Vous recherchez l’ensemble de conditions requises de l’API *propres à l’hôte* ? Consultez les ensembles de conditions requises d’API suivants :
 
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md) (ExcelApi)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md) (WordApi)
- [Ensembles de conditions requises de l’API JavaScript pour OneNote](onenote-api-requirement-sets.md) (OneNoteApi)
- [Présentation de l’ensemble de conditions requises pour les API Outlook](outlook-api-requirement-sets.md) (MailBox)

> [!IMPORTANT]
> Désormais, nous ne recommandons plus la création et l'utilisation des applications web Access et des bases de données dans SharePoint. Comme alternative, nous vous recommandons d’utiliser [Microsoft PowerApps](https://powerapps.microsoft.com/) pour créer des solutions sans code pour les appareils mobiles et web.

## <a name="common-api-requirement-sets"></a>Ensembles de conditions requises des API communes

Le tableau suivant répertorie les ensembles de conditions requises communes, les méthodes de chaque ensemble et les applications hôtes Office qui les prennent en charge. Tous ces ensembles de conditions requises d’API sont à la version 1.1.

|**Ensemble de conditions requises**|**Hôte Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|:-----|
| ActiveView | PowerPoint<br>PowerPoint Online<br>PowerPoint pour iPad<br>PowerPoint pour Mac|Document.getActiveViewAsync|
| AddInCommands | Voir [Ensembles de conditions requises pour les commandes de complément](add-in-commands-requirement-sets.md). | |
| BindingEvents  | Applications web Access<br>Excel<br>Excel Online<br>Excel pour iPad<br>Excel pour Mac<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | Excel<br>Excel Online<br>Excel pour iPad<br>Excel pour Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint pour iPad<br>PowerPoint pour Mac<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|Prend en charge la sortie au format Office Open XML (OOXML) sous la forme d’un tableau d’octets<br>(Office.FileType.Compressed) lorsque vous utilisez la méthode Document.getFileAsync.|
| CustomXmlParts    | Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DialogApi | Voir [l'ensembles de conditions requises de l’API de boîte de dialogue](dialog-api-requirement-sets.md) | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |
| DocumentEvents    | Excel<br>Excel Online<br>Excel pour iPad<br>Excel pour Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint pour iPad<br>PowerPoint pour Mac<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| Fichier  | Excel<br>Excel Online<br>Excel pour iPad<br>Excel pour Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint pour iPad<br>PowerPoint pour Mac<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | OneNote Online<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|Prend en charge le forçage au format HTML (Office.CoercionType.Html) lors de la lecture et de l’écriture des données à l’aide des méthodes Document.getSelectedDataAsync,<br>Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| IdentityAPI | Voir [Ensembles de conditions requises de l’API d’identité](identity-api-requirement-sets.md) | Auth.getAccessTokenAsync |
| ImageCoercion | Excel<br>Excel pour iPad<br>Excel pour Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint pour iPad<br>PowerPoint pour Mac<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|Prise en charge de la conversion en une image (Office.CoercionType.Image) lors de l’écriture des données à l’aide de la méthode Document.setSelectedDataAsync.|
| Boîte aux lettres   |Outlook pour Windows<br>Outlook pour le web<br>Outlook pour Android<br>Outlook pour Mac<br>Application web Outlook |Voir [Présentation de l’ensemble de conditions requises pour les API Outlook](outlook-api-requirement-sets.md).|
| MatrixBindings    | Excel<br>Excel Online<br>Excel pour iPad<br>Excel pour Mac<br>Word<br>Word Online<br>Word pour iPad<br>Word pour Mac|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | Excel<br>Excel Online<br>Excel pour iPad<br>Excel pour Mac<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|Prise en charge du forçage de type sur la structure de données (Office.CoercionType.Matrix) « matrice » (tableau de tableaux) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| OoxmlCoercion | Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|Prise en charge du forçage de type au format Open Office XML (OOXML) (Office.CoercionType.Ooxml) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| PartialTableBindings  | Applications web Access||
| PdfFile   | Excel pour Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint pour iPad<br>PowerPoint pour Mac<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|Prend en charge la sortie au format PDF (Office.FileType.Pdf)<br>lorsque vous utilisez la méthode Document.getFileAsync.|
| Sélection | Excel<br>Excel Online<br>Excel pour iPad<br>Excel pour Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint pour iPad<br>PowerPoint pour Mac<br>Projet<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Paramètres  | Applications web Access<br>Excel<br>Excel Online<br>Excel pour iPad<br>Excel pour Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint pour iPad<br>PowerPoint pour Mac<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|Settings.Get<br>Settings.Remove<br>Settings.saveAsync<br>Settings.Set|
| TableBindings | Applications web Access<br>Excel<br>Excel Online<br>Excel pour iPad<br>Excel pour Mac<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | Applications web Access<br>Excel<br>Excel Online<br>Excel pour iPad<br>Excel pour Mac<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|Prise en charge du forçage de type sur la structure de données « tableau » (Office.CoercionType.Table) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| TextBindings  | Excel<br>Excel Online<br>Excel pour iPad<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | Excel<br>Excel Online<br>Excel pour iPad<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint pour iPad<br>PowerPoint pour Mac<br>Projet<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|Prise en charge du forçage de type au format texte (Office.CoercionType.Text) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| TextFile  | Word 2013 et versions ultérieures<br>Word 2016 pour Mac et versions ultérieures<br>Word Online<br>Word pour iPad|Prise en charge de sortie au format texte (Office.FileType.Text) lors de l’utilisation de la méthode Document.getFileAsync.|

## <a name="methods-that-arent-part-of-a-requirement-set"></a>Méthodes qui ne font pas partie d’un ensemble de conditions requises

Les méthodes suivantes de l’API JavaScript pour Office ne font pas partie d’un ensemble de conditions requises. Si votre complément requiert l’une de ces méthodes, utilisez les  **méthodes** et éléments de la **méthode** dans le manifeste du complément pour déclarer qu’ils sont requis ou effectuer la vérification de runtime à l’aide d'une `if` instruction. Pour plus d’informations, consultez la rubrique [Spécifier les hôtes Office et les conditions requises d’API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).

|**Nom de la méthode**|**Support de l'hôte Office**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Applications web Access, Excel, Excel Online et Excel pour iPad|
|Document.getFilePropertiesAsync|Excel, Excel Online, Excel pour iPad, Excel pour Mac, PowerPoint, PowerPoint Online, PowerPoint pour iPad, PowerPoint pour Mac, Word, Word Online, Word pour iPad et Word pour Mac|
|Document.getProjectFieldAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getResourceFieldAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getSelectedViewAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getTaskAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getTaskFieldAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.goToByIdAsync|Excel, Excel Online, Excel pour iPad, Excel pour Mac, PowerPoint, PowerPoint Online, PowerPoint pour iPad, PowerPoint pour Mac, Word, Word Online, Word pour iPad et Word pour Mac|
|Settings.addHandlerAsync|Applications web Access, Excel, Excel Online, PowerPoint, PowerPoint Online, Word et Word Online|
|Settings.refreshAsync|Applications web Access, Excel, Excel Online, PowerPoint, PowerPoint Online, Word et Word Online|
|Settings.removeHandlerAsync|Applications web Access, Excel, Excel Online, PowerPoint, PowerPoint Online, Word et Word Online|
|TableBinding.clearFormatsAsync|Excel, Excel Online et Excel pour Mac|
|TableBinding.setFormatsAsync|Excel, Excel Online, Excel pour iPad et Excel pour Mac|
|TableBinding.setTableOptionsAsync|Excel, Excel Online, Excel pour iPad et Excel pour Mac|

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Spécification des conditions requises d’hôtes Office et d’API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifeste XML des compléments Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
