---
title: Ensembles de conditions requises des API communes pour Office
description: ''
ms.date: 07/17/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: b37adca116c60b465e11858cb813e9a7f9247ed3
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950550"
---
# <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Vous avez besoin d’informations sur l’endroit où les compléments sont pris en charge par l’hôte Office ? Consultez la rubrique [Disponibilité des compléments Office sur les plateformes et les hôtes](/office/dev/add-ins/overview/office-add-in-availability).

Vous recherchez l’ensemble de conditions requises de l’API *propres à l’hôte* ? Reportez-vous aux ensembles de conditions requises des API suivants :

- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md) (ExcelApi)
- [Ensembles de conditions requises de l’API JavaScript pour Word](word-api-requirement-sets.md) (WordApi)
- [Ensembles de conditions requises de l’API JavaScript pour OneNote](onenote-api-requirement-sets.md) (OneNoteApi)
- [Ensembles de conditions requises de l’API JavaScript pour PowerPoint](powerpoint-api-requirement-sets.md) (PowerPointApi)
- [Présentation de l’ensemble de conditions requises pour les API Outlook](outlook-api-requirement-sets.md) (Mailbox)

> [!IMPORTANT]
> Nous ne vous recommandons plus de créer et d’utiliser les bases de données et les applications web Access dans SharePoint. Nous vous recommandons plutôt d’utiliser [Microsoft PowerApps](https://powerapps.microsoft.com/) pour créer des solutions professionnelles sans code pour des appareils mobiles et web.

## <a name="common-api-requirement-sets"></a>Ensembles de conditions requises des API communes

Les sections suivantes répertorient les ensembles communs de conditions requises, les méthodes de chaque ensemble et les applications hôtes Office qui les prennent en charge. Tous ces ensembles de conditions requises d’API sont à la version 1.1., sauf indication contraire.

### <a name="activeview"></a>ActiveView

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| PowerPoint sur Windows<br>PowerPoint sur le web<br>PowerPoint sur iPad<br>PowerPoint sur Mac|Document.getActiveViewAsync|

---

### <a name="addincommands"></a>AddInCommands

Consultez la rubrique [Exigences relatives aux commandes de complément](add-in-commands-requirement-sets.md).

---

### <a name="bindingevents"></a>BindingEvents

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Applications web Access<br>Excel sur Windows<br>Excel sur le web<br>Excel sur iPad<br>Excel sur Mac<br>Word 2013 ou version ultérieure sur Windows<br>Word 2016 ou version ultérieure sur Mac<br>Word sur le web<br>Word sur iPad|Binding.addHandlerAsync<br>Binding.removeHandlerAsync|

---

### <a name="compressedfile"></a>CompressedFile

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Excel sur Windows<br>Excel sur le web<br>Excel sur Mac<br>PowerPoint sur Windows<br>PowerPoint sur le web<br>PowerPoint sur iPad<br>PowerPoint sur Mac<br>Word 2013 ou version ultérieure sur Windows<br>Word 2016 ou version ultérieure sur Mac<br>Word sur le web<br>Word sur iPad|Prend en charge la sortie au format Office Open XML (OOXML) sous la forme d’un tableau d’octets<br>(Office.FileType.Compressed) lorsque vous utilisez la méthode Document.getFileAsync.|

---

### <a name="customxmlparts"></a>CustomXmlParts

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Word 2013 ou version ultérieure sur Windows<br>Word 2016 ou version ultérieure sur Mac<br>Word sur le web<br>Word sur iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getTextAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setTextAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|

---

### <a name="dialogapi"></a>DialogApi

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Consultez la rubrique [Ensembles de conditions requises de l’API de boîte de dialogue](dialog-api-requirement-sets.md). | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |

---

### <a name="documentevents"></a>DocumentEvents

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Excel sur Windows<br>Excel sur le web<br>Excel sur iPad<br>Excel sur Mac<br>OneNote sur le web<br>PowerPoint sur Windows<br>PowerPoint sur le web<br>PowerPoint sur iPad<br>PowerPoint sur Mac<br>Word 2013 ou version ultérieure sur Windows<br>Word 2016 ou version ultérieure sur Mac<br>Word sur le web<br>Word sur iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|

---

### <a name="file"></a>File

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Excel sur Windows<br>Excel sur le web<br>Excel sur iPad<br>Excel sur Mac<br>PowerPoint sur Windows<br>PowerPoint sur le web<br>PowerPoint sur iPad<br>PowerPoint sur Mac<br>Word 2013 ou version ultérieure sur Windows<br>Word 2016 ou version ultérieure sur Mac<br>Word sur le web<br>Word sur iPad|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|

---

### <a name="htmlcoercion"></a>HtmlCoercion

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| OneNote sur le web<br>Word 2013 ou version ultérieure sur Windows<br>Word 2016 ou version ultérieure sur Mac<br>Word sur le web<br>Word sur iPad|Prise en charge du forçage de type au format HTML (Office.CoercionType.Html) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|

---

### <a name="identityapi"></a>IdentityAPI

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Consultez la rubrique [Ensembles de conditions requises de l’API d’identité](identity-api-requirement-sets.md). | Auth.getAccessToken |

---

### <a name="imagecoercion"></a>ImageCoercion

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Voir [Ensembles de conditions requises de coercition d’image](image-coercion-requirement-sets.md). | Méthode Document.setSelectedDataAsync|

---

### <a name="mailbox"></a>Boîte aux lettres

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
|Outlook sur Windows<br>Outlook sur le web<br>Outlook sur Android<br>Outlook sur Mac<br>Outlook sur iOS|Voir [Présentation de l’ensemble de conditions requises pour les API Outlook](outlook-api-requirement-sets.md).|

---

### <a name="matrixbindings"></a>MatrixBindings

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Excel sur Windows<br>Excel sur le web<br>Excel sur iPad<br>Excel sur Mac<br>Word sur Windows<br>Word sur le web<br>Word sur iPad<br>Word sur Mac|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="matrixcoercion"></a>MatrixCoercion

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Excel sur Windows<br>Excel sur le web<br>Excel sur iPad<br>Excel sur Mac<br>Word 2013 ou version ultérieure sur Windows<br>Word 2016 ou version ultérieure sur Mac<br>Word sur le web<br>Word sur iPad|Prise en charge du forçage de type sur la structure de données (Office.CoercionType.Matrix) « matrice » (tableau de tableaux) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|

---

### <a name="ooxmlcoercion"></a>OoxmlCoercion

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Word 2013 ou version ultérieure sur Windows<br>Word 2016 ou version ultérieure sur Mac<br>Word sur le web<br>Word sur iPad|Prise en charge du forçage de type au format Open Office XML (OOXML) (Office.CoercionType.Ooxml) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|

---

### <a name="partialtablebindings"></a>PartialTableBindings

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Applications web Access||

---

### <a name="pdffile"></a>PdfFile

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Excel sur Mac<br>PowerPoint sur Windows<br>PowerPoint sur le web<br>PowerPoint sur iPad<br>PowerPoint sur Mac<br>Word 2013 ou version ultérieure sur Windows<br>Word 2016 ou version ultérieure sur Mac<br>Word sur le web<br>Word sur iPad|Prend en charge la sortie au format PDF (Office.FileType.Pdf)<br>lorsque vous utilisez la méthode Document.getFileAsync.|

---

### <a name="selection"></a>Selection

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Excel sur Windows<br>Excel sur le web<br>Excel sur iPad<br>Excel sur Mac<br>PowerPoint sur Windows<br>PowerPoint sur le web<br>PowerPoint sur iPad<br>PowerPoint sur Mac<br>Project sur Windows<br>Word 2013 ou version ultérieure sur Windows<br>Word 2016 ou version ultérieure sur Mac<br>Word sur le web<br>Word sur iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|

---

### <a name="settings"></a>Settings

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Applications web Access<br>Excel sur Windows<br>Excel sur le web<br>Excel sur iPad<br>Excel sur Mac<br>OneNote sur le web<br>PowerPoint sur Windows<br>PowerPoint sur le web<br>PowerPoint sur iPad<br>PowerPoint sur Mac<br>Word 2013 ou version ultérieure sur Windows<br>Word 2016 ou version ultérieure sur Mac<br>Word sur le web<br>Word sur iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|

---

### <a name="tablebindings"></a>TableBindings

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Applications web Access<br>Excel sur Windows<br>Excel sur le web<br>Excel sur iPad<br>Excel sur Mac<br>Word 2013 ou version ultérieure sur Windows<br>Word 2016 ou version ultérieure sur Mac<br>Word sur le web<br>Word sur iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.addColumnsAsync<br>Binding.addRowsAsync<br>Binding.deleteAllDataValuesAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="tablecoercion"></a>TableCoercion

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Applications web Access<br>Excel sur Windows<br>Excel sur le web<br>Excel sur iPad<br>Excel sur Mac<br>Word 2013 ou version ultérieure sur Windows<br>Word 2016 ou version ultérieure sur Mac<br>Word sur le web<br>Word sur iPad|Prise en charge du forçage de type sur la structure de données « tableau » (Office.CoercionType.Table) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|

---

### <a name="textbindings"></a>TextBindings

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Excel sur Windows<br>Excel sur le web<br>Excel sur iPad<br>Excel sur Mac<br>Word 2013 ou version ultérieure et Windows<br>Word 2016 ou version ultérieure sur Mac<br>Word sur le web<br>Word sur iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="textcoercion"></a>TextCoercion

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Excel sur Windows<br>Excel sur le web<br>Excel sur iPad<br>OneNote sur le web<br>PowerPoint sur Windows<br>PowerPoint sur le web<br>PowerPoint sur iPad<br>PowerPoint sur Mac<br>Project sur Windows<br>Word 2013 ou version ultérieure sur Windows<br>Word 2016 ou version ultérieure sur Mac<br>Word sur le web<br>Word sur iPad|Prise en charge du forçage de type au format texte (Office.CoercionType.Text) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|

---

### <a name="textfile"></a>TextFile

|**Hôtes Office**|**Méthodes dans l’ensemble**|
|:-----|:-----|
| Word 2013 ou version ultérieure sur Windows<br>Word 2016 ou version ultérieure sur Mac<br>Word sur le web<br>Word sur iPad|Prise en charge de sortie au format texte (Office.FileType.Text) lors de l’utilisation de la méthode Document.getFileAsync.|

---

## <a name="methods-that-arent-part-of-a-requirement-set"></a>Méthodes qui ne font pas partie d’un ensemble de conditions requises

Les méthodes suivantes dans l’API JavaScript pour Office ne font pas partie d’un ensemble de conditions requises. Si votre complément requiert l’une de ces méthodes, utilisez les éléments **Methods** et **Method** dans le manifeste du complément afin de déclarer qu’ils sont requis ou effectuez la vérification de l’exécution à l’aide d’une instruction `if`. Pour plus d’informations, consultez la rubrique [Spécifier les hôtes Office et la configuration requise d’API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).

|**Nom de la méthode**|**Prise en charge des hôtes Office**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Accès aux applications web, Excel sur Windows, Excel sur le web, Excel sur iPad et Excel sur Mac|
|Document.getFilePropertiesAsync|Excel sur Windows, Excel sur le web, Excel sur iPad, Excel sur Mac, PowerPoint sur Windows, PowerPoint sur le web, PowerPoint sur iPad, PowerPoint sur Mac, Word sur Windows, Word sur le web, Word sur iPad et Word sur Mac|
|Document.getProjectFieldAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getResourceFieldAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getSelectedViewAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getTaskAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getTaskFieldAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.goToByIdAsync|Excel sur Windows, Excel sur le web, Excel sur iPad, Excel sur Mac, PowerPoint sur Windows, PowerPoint sur le web, PowerPoint sur iPad, PowerPoint sur Mac, Word sur Windows, Word sur le web, Word sur iPad et Word sur Mac|
|Settings.addHandlerAsync|Accès aux applications web et Excel sur le web|
|Settings.refreshAsync|Accès aux applications web, Excel sur Windows, Excel sur le web, PowerPoint sur Windows, PowerPoint sur le web, Word et Word sur le web|
|Settings.removeHandlerAsync|Accès aux applications web et Excel sur le web|
|TableBinding.clearFormatsAsync|Excel sur Windows, Excel sur le web, Excel sur iPad et Excel sur le web|
|TableBinding.setFormatsAsync|Excel sur Windows, Excel sur le web, Excel sur iPad et Excel sur le web|
|TableBinding.setTableOptionsAsync|Excel sur Windows, Excel sur le web, Excel sur iPad et Excel sur le web|

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Spécification des exigences en matière d’hôtes Office et d’API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifeste XML des compléments Office](/office/dev/add-ins/develop/add-in-manifests)
