---
title: Application cliente Office et disponibilité de la plate-forme pour les compléments Office
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 07/13/2021
localization_priority: Priority
ms.openlocfilehash: 983d843cee4c0df29fc16445b306a47fabf240c6c76a091b8e319e5a705ee9c2
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57083911"
---
# <a name="office-client-application-and-platform-availability-for-office-add-ins"></a>Application cliente Office et disponibilité de la plate-forme pour les compléments Office

Pour fonctionner comme prévu, votre complément Office peut dépendre d'une application Office spécifique, d'un ensemble de conditions requises, d'un membre de l’API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API courantes qui sont actuellement prises en charge pour chaque application Office.

<br>

|<a href="#excel"><img src="../images/index/logo-excel.svg" alt="Excel" width="48" /><br><span>Excel</span></a>|<a href="#onenote"><img src="../images/index/logo-onenote.svg" alt="OneNote" width="48" /><br><span>OneNote</span></a>|<a href="#outlook"><img src="../images/index/logo-outlook.svg" alt="Outlook" width="48" /><br><span>Outlook</span></a>|<a href="#powerpoint"><img src="../images/index/logo-powerpoint.svg" alt="PowerPoint" width="48" /><br><span>PowerPoint</span></a>|<a href="#project"><img src="../images/index/logo-project-server.svg" alt="Project" width="48" /><br><span>Project</span></a>|<a href="#word"><img src="../images/index/logo-word.svg" alt="Word" width="48" /><br><span>Word</span></a>|
|:---:|:---:|:---:|:---:|:---:|:---:|

> [!NOTE]
> La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune. Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also). Les compléments Office ne sont peut-être pas pris en charge sur tous les services membres du [Programme partenaire de stockage cloud Office](https://developer.microsoft.com/office/cloud-storage-partner-program), qui permet à l’intégration d’Office sur le web d’utiliser des documents Office dans le cadre de leur offre de service. Pour plus d’informations, contactez le service membre.

## <a name="excel"></a>Excel

<table style="width:80%">
  <tr>
    <th style="width:10%">Plate-forme</th>
    <th style="width:10%">Points d’extension</th>
    <th style="width:20%">Ensembles de conditions requises de l’API</th>
    <th style="width:40%"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>API communes</b></a></th>
  </tr>
  <tr>
    <td>Office sur le web</td>
    <td>
      - Volet Office<br>
      - Contenu<br>
      - CustomFunctions<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a><br>
      - <a href="../reference/requirement-sets/excel-api-online-requirement-set.md">ExcelApiOnline</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonAPI 1.1</a><br>
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office pour Windows<br>(connecté à un abonnement Microsoft 365)</td>
    <td>
      - Volet Office<br>
      - Contenu<br>
      - CustomFunctions<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a><br>
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonAPI 1.1</a><br>
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 sur Windows<br>(achat définitif)</td>
    <td>
      - Volet Office<br>
      - Contenu<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 sur Windows<br>(achat définitif)</td>
    <td>
      - Volet Office<br>
      - Contenu </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2013 sur Windows<br>(achat définitif)</td>
    <td>
      - Volet Office<br>
      - Contenu </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office sur iPad<br>(connecté à un abonnement Microsoft 365)</td>
    <td>
      - Volet Office<br>
      - Contenu </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office sur Mac<br>(connecté à un abonnement Microsoft 365)</td>
    <td>
      - Volet Office<br>
      - Contenu<br>
      - CustomFunctions<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a><br>
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonAPI 1.1</a><br>
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 sur Mac<br>(achat définitif)</td>
    <td>
      - Volet Office<br>
      - Contenu<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 sur Mac<br>(achat définitif)</td>
    <td>
      - Volet Office<br>
      - Contenu </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

*&ast; : ajouté avec les mises à jour après la publication.*

## <a name="custom-functions-excel-only"></a>Fonctions personnalisées (Excel seulement)

<table style="width:80%">
  <tr>
    <th>Plateforme</th>
    <th>Points d’extension</th>
    <th>Ensembles de conditions requises de l’API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>API communes</b></a></th>
  </tr>
  <tr>
    <td>Office sur le web</td>
    <td>- CustomFunctions</td>
    <td>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </td>
    <td></td>
  </tr>
  <tr>
    <td>Office pour Windows<br>(connecté à un abonnement Microsoft 365)</td>
    <td>- CustomFunctions</td>
    <td>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </td>
    <td></td>
  </tr>
  <tr>
    <td>Office sur Mac<br>(connecté à un abonnement Microsoft 365)</td>
    <td>- CustomFunctions</td>
    <td>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </td>
    <td></td>
  </tr>
</table>

## <a name="outlook"></a>Outlook

<table style="width:80%">
  <tr>
    <th>Plate-forme</th>
    <th>Points d’extension</th>
    <th>Ensembles de conditions requises de l’API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>API communes</b></a></th>
  </tr>
  <tr>
    <td>Office sur le web<br>(moderne)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Mailbox 1.9</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Mailbox 1.10</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office sur le web<br>(classique)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office pour Windows<br>(connecté à un abonnement Microsoft 365)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a><br>
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Mailbox 1.9</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Mailbox 1.10</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office 2019 sur Windows<br>(achat définitif)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a><br>
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a>
    </td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office 2016 sur Windows<br>(achat définitif)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a><br>
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><sup>2</sup>
    </td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office 2013 sur Windows<br>(achat définitif)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><sup>2</sup><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><sup>2</sup>
    </td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office sur iOS<br>(connecté à un abonnement Microsoft 365)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Message lu</a><br>
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Organisateur de rendez-vous (composer) : réunion en ligne</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a>
    </td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office sur Mac<br>(interface utilisateur actuelle<br>connectée à un abonnement Microsoft 365)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office sur Mac<br>(nouvelle interface utilisateur (aperçu)<sup>3</sup>,<br>connectée à un abonnement Microsoft 365)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office 2019 sur Mac<br>(achat définitif)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office 2016 sur Mac<br>(achat définitif)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message lu</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office sur Android<br>(connecté à un abonnement Microsoft 365)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Message lu</a><br>
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Organisateur de rendez-vous (composer) : réunion en ligne</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a>
    </td>
    <td>Non disponible</td>
  </tr>
</table>

> [!NOTE]
> <sup>1</sup> Pour nécessiter le jeu d'API d'identité 1.3 dans votre code additionnel, vérifiez s'il est pris en charge en appelant `isSetSupported('IdentityAPI', '1.3')`. Le déclarer dans le manifeste de votre macro complémentaire n'est pas pris en charge. Vous pouvez également déterminer si l’API est prise en charge en vérifiant qu’elle n’est pas `undefined`. Pour plus d’informations, consultez [Utilisation des API d’un ensemble de conditions requises ultérieure](../reference/requirement-sets/outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).
>
> <sup>2</sup> Ajouté avec les mises à jour après la publication.
>
> <sup>3</sup> prise en charge de la préversion pour le nouvel Outlook sur Mac est disponible dans la version 16.38.506. Pour plus d’informations, consultez la section [Prise en charge du macro complémentaire dans Outlook sur le nouvel interface d’utilisateur Mac](../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview).

> [!IMPORTANT]
> La prise en charge du client pour un ensemble de conditions requises peut être limitée par la prise en charge d’Exchange Server. Consultez [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) pour plus d’informations sur les ensembles de conditions requises pris en charge par Exchange Server et le client Outlook.

<br/>

## <a name="word"></a>Word

<table style="width:80%">
  <tr>
    <th>Plate-forme</th>
    <th>Points d’extension</th>
    <th>Ensembles de conditions requises de l’API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>API communes</b></a></th>
  </tr>
  <tr>
    <td>Office sur le web</td>
    <td>
      - Volet Office<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office pour Windows<br>(connecté à un abonnement Microsoft 365)</td>
    <td>
      - Volet Office<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 sur Windows<br>(achat définitif)</td>
    <td>
      - Volet Office<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 sur Windows<br>(achat définitif)</td>
    <td>- Volet Office</td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2013 sur Windows<br>(achat définitif)</td>
    <td>- Volet Office</td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office sur iPad<br>(connecté à un abonnement Microsoft 365)</td>
    <td>- Volet Office</td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office sur Mac<br>(connecté à un abonnement Microsoft 365)</td>
    <td>
      - Volet Office<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 sur Mac<br>(achat définitif)</td>
    <td>
      - Volet Office<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 sur Mac<br>(achat définitif)</td>
    <td>- Volet Office</td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
</table>

*&ast; : ajouté avec les mises à jour après la publication.*

<br/>

## <a name="powerpoint"></a>PowerPoint

<table style="width:80%">
  <tr>
    <th>Plate-forme</th>
    <th>Points d’extension</th>
    <th>Ensembles de conditions requises de l’API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>API communes</b></a></th>
  </tr>
  <tr>
    <td>Office sur le web</td>
    <td>
      - Contenu<br>
      - Volet Office<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a><br>
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office pour Windows<br>(connecté à un abonnement Microsoft 365)</td>
    <td>
      - Contenu<br>
      - Volet Office<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a><br>
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 sur Windows<br>(achat définitif)</td>
    <td>
      - Contenu<br>
      - Volet Office<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 sur Windows<br>(achat définitif)</td>
    <td>
      - Contenu<br>
      - Volet Office </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2013 sur Windows<br>(achat définitif)</td>
    <td>
      - Contenu<br>
      - Volet Office </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office sur iPad<br>(connecté à un abonnement Microsoft 365)</td>
    <td>
      - Contenu<br>
      - Volet Office </td>
    <td>
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a><br>
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office sur Mac<br>(connecté à un abonnement Microsoft 365)</td>
    <td>
      - Contenu<br>
      - Volet Office<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a><br>
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 sur Mac<br>(achat définitif)</td>
    <td>
      - Contenu<br>
      - Volet Office<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 sur Mac<br>(achat définitif)</td>
    <td>
      - Contenu<br>
      - Volet Office </td>
    <td>
       - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

*&ast; : ajouté avec les mises à jour après la publication.*

<br/>

## <a name="onenote"></a>OneNote

<table style="width:80%">
  <tr>
    <th>Plate-forme</th>
    <th>Points d’extension</th>
    <th>Ensembles de conditions requises de l’API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>API communes</b></a></th>
  </tr>
  <tr>
    <td>Office sur le web</td>
    <td>
      - Contenu<br>
      - Volet Office<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Commandes de complément</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/onenote-api-requirement-sets.md">OneNoteApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Paramètres</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

<br/>

## <a name="project"></a>Project

<table style="width:80%">
  <tr>
    <th>Plateforme</th>
    <th>Points d’extension</th>
    <th>Ensembles de conditions requises de l’API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>API communes</b></a></th>
  </tr>
  <tr>
    <td>Office 2019 sur Windows<br>(achat définitif)</td>
    <td>- Volet Office</td>
    <td>- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 sur Windows<br>(achat définitif)</td>
    <td>- Volet Office</td>
    <td>- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2013 sur Windows<br>(achat définitif)</td>
    <td>- Volet Office</td>
    <td>- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

<br/>

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la plateforme des compléments Office](office-add-ins.md)
- [Versions d’Office et ensembles de conditions requises](../develop/office-versions-and-requirement-sets.md)
- [Ensembles de conditions requises des API communes](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Ensembles de conditions requises concernant les commandes de complément](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [Documentation de référence de l’API](../reference/javascript-api-for-office.md)
- [Historique des mises à jour de Microsoft 365 Apps](/officeupdates/update-history-office365-proplus-by-date)
- [Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)](/officeupdates/update-history-office-2019)
- [Historique des mises à jour d’Office 2013 (Démarrer en un clic)](/officeupdates/update-history-office-2013)
- [Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)](/officeupdates/office-updates-msi)
- [Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)](/officeupdates/outlook-updates-msi)
- [Historique des mises à jour d’Office pour Mac](/officeupdates/update-history-office-for-mac)
- [Développement de compléments Office](../develop/develop-overview.md)
