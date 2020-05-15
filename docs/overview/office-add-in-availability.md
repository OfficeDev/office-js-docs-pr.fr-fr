---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 05/11/2020
localization_priority: Priority
ms.openlocfilehash: 36c6bc6b6348ac988049f9a50127f6dd2f94bf37
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217822"
---
# <a name="office-add-in-host-and-platform-availability"></a>Disponibilité des compléments Office sur les plateformes et les hôtes

Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.

> [!NOTE]
> La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune. Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).

## <a name="excel"></a>Excel

<table style="width:80%">
  <tr>
    <th style="width:10%">Plate-forme</th>
    <th style="width:10%">Points d’extension</th>
    <th style="width:20%">Ensembles de conditions requises de l’API</th>
    <th style="width:40%"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></th>
  </tr>
  <tr>
    <td>Office sur le web</td>
    <td> - Volet Office<br>
        - Contenu<br>
        - Fonctions personnalisées<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
        - BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office pour Windows<br>(connecté à l’abonnement Office 365)</td>
    <td> - Volet Office<br>
        - Contenu<br>
        - Fonctions personnalisées<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td>
        - BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2019 sur Windows<br>(achat définitif)</td>
    <td>- Volet Office<br>
        - Contenu<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 sur Windows<br>(achat définitif)</td>
    <td>- Volet Office<br>
        - Contenu</td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 sur Windows<br>(achat définitif)</td>
    <td>
        - Volet Office<br>
        - Contenu</td>
    <td>  - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>
        - BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office sur iPad<br>(connecté à l’abonnement Office 365)</td>
    <td>- Volet Office<br>
        - Contenu</td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>- BindingEvents<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office sur Mac<br>(connecté à l’abonnement Office 365)</td>
    <td>- Volet Office<br>
        - Contenu<br>
        - Fonctions personnalisées<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - PdfFile<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2019 sur Mac<br>(achat définitif)</td>
    <td>- Volet Office<br>
        - Contenu<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - PdfFile<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 sur Mac<br>(achat définitif)</td>
    <td>- Volet Office<br>
        - Contenu</td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - PdfFile<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
</table>

*&ast; : ajouté avec les mises à jour après la publication.*

## <a name="custom-functions-excel-only"></a>Fonctions personnalisées (Excel seulement)

<table style="width:80%">
  <tr>
    <th style="width:10%">Plateforme</th>
    <th style="width:10%">Points d’extension</th>
    <th style="width:20%">Ensembles de conditions requises de l’API</th>
    <th style="width:40%"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></th>
  </tr>
  <tr>
    <td>Office sur le web</td>
    <td>
        - Fonctions personnalisées</td>
    <td>
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></td>
    <td>
    </td>
  </tr>
  <tr>
    <td>Office pour Windows<br>(connecté à l’abonnement Office 365)</td>
    <td>
        - Fonctions personnalisées</td>
    <td>
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></td>
    <td>
    </td>
  </tr>
  <tr>
    <td>Office pour Mac<br>(connecté à Office 365)</td>
    <td>
        - Fonctions personnalisées</td>
    <td>
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a>Outlook

<table style="width:80%">
  <tr>
    <th>Plate-forme</th>
    <th>Points d’extension</th>
    <th>Ensembles de conditions requises de l’API</th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></th>
  </tr>
  <tr>
    <td>Office sur le web<br>(moderne)</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office sur le web<br>(classique)</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office pour Windows<br>(connecté à l’abonnement Office 365)</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office 2019 sur Windows<br>(achat définitif)</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office 2016 sur Windows<br>(achat définitif)</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office 2013 sur Windows<br>(achat définitif)</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*<br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office sur iOS<br>(connecté à l’abonnement Office 365)</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message lu</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office sur Mac<br>(connecté à l’abonnement Office 365)</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office 2019 sur Mac<br>(achat définitif)</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office 2016 sur Mac<br>(achat définitif)</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message lu</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composer un message</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participant au rendez-vous (lecture)</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organisateur de rendez-vous (composer)</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></td>
    <td>Non disponible</td>
  </tr>
  <tr>
    <td>Office sur Android<br>(connecté à l’abonnement Office 365)</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message lu</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Organisateur de rendez-vous (composer) : réunion en ligne</a> (aperçu)<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></td>
    <td>Non disponible</td>
  </tr>
</table>

*&ast; : ajouté avec les mises à jour après la publication.*

> [!IMPORTANT]
> La prise en charge du client pour un ensemble de conditions requises peut être limitée par la prise en charge d’Exchange Server. Consultez [Ensembles de conditions requises de l’API JavaScript pour Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) pour plus d’informations sur les ensembles de conditions requises pris en charge par Exchange Server et le client Outlook.

<br/>

## <a name="word"></a>Word

<table style="width:80%">
  <tr>
    <th>Plate-forme</th>
    <th>Points d’extension</th>
    <th>Ensembles de conditions requises de l’API</th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></th>
  </tr>
  <tr>
    <td>Office sur le web</td>
    <td> - Volet Office<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td> - BindingEvents<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile</td>
  </tr>
  <tr>
    <td>Office pour Windows<br>(connecté à l’abonnement Office 365)</td>
    <td> - Volet Office<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Office 2019 sur Windows<br>(achat définitif)</td>
    <td> - Volet des tâches<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Office 2016 sur Windows<br>(achat définitif)</td>
    <td> - Volet Office</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Office 2013 sur Windows<br>(achat définitif)</td>
    <td> - Volet Office</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile</td>
  </tr>
  <tr>
    <td>Office sur iPad<br>(connecté à l’abonnement Office 365)</td>
    <td> - Volet Office</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Office sur Mac<br>(connecté à l’abonnement Office 365)</td>
    <td> - Volet Office<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Office 2019 sur Mac<br>(achat définitif)</td>
    <td> - Volet des tâches<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Office 2016 sur Mac<br>(achat définitif)</td>
    <td> - Volet Office</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
</table>

*&ast; : ajouté avec les mises à jour après la publication.*

<br/>

## <a name="powerpoint"></a>PowerPoint

<table style="width:80%">
  <tr>
    <th>Plate-forme</th>
    <th>Points d’extension</th>
    <th>Ensembles de conditions requises de l’API</th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></th>
  </tr>
  <tr>
    <td>Office sur le web</td>
    <td> - Contenu<br>
         - Volet Office<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office pour Windows<br>(connecté à l’abonnement Office 365)</td>
    <td> - Contenu<br>
         - Volet Office<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2019 sur Windows<br>(achat définitif)</td>
    <td> - Contenu<br>
         - Volet Office<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 sur Windows<br>(achat définitif)</td>
    <td> - Contenu<br>
         - Volet Office</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 sur Windows<br>(achat définitif)</td>
    <td> - Contenu<br>
         - Volet Office<br>
    </td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office sur iPad<br>(connecté à l’abonnement Office 365)</td>
    <td> - Contenu<br>
         - Volet Office</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office sur Mac<br>(connecté à l’abonnement Office 365)</td>
    <td> - Contenu<br>
         - Volet Office<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2019 sur Mac<br>(achat définitif)</td>
    <td> - Contenu<br>
         - Volet Office<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 sur Mac<br>(achat définitif)</td>
    <td> - Contenu<br>
         - Volet Office</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
</table>

*&ast; : ajouté avec les mises à jour après la publication.*

<br/>

## <a name="onenote"></a>OneNote

<table style="width:80%">
  <tr>
    <th>Plate-forme</th>
    <th>Points d’extension</th>
    <th>Ensembles de conditions requises de l’API</th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></th>
  </tr>
  <tr>
    <td>Office sur le web</td>
    <td> - Contenu<br>
         - Volet Office<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - DocumentEvents<br>
         - HtmlCoercion<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
</table>

<br/>

## <a name="project"></a>Projet

<table style="width:80%">
  <tr>
    <th>Plateforme</th>
    <th>Points d’extension</th>
    <th>Ensembles de conditions requises de l’API</th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></th>
  </tr>
  <tr>
    <td>Office 2019 sur Windows<br>(achat définitif)</td>
    <td> - Volet Office</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - Selection<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 sur Windows<br>(achat définitif)</td>
    <td> - Volet Office</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - Selection<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 sur Windows<br>(achat définitif)</td>
    <td> - Volet Office</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - Selection<br>
         - TextCoercion</td>
  </tr>
</table>

<br/>

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la plateforme des compléments Office](office-add-ins.md)
- [Versions d’Office et ensembles de conditions requises](../develop/office-versions-and-requirement-sets.md)
- [Ensembles de conditions requises des API communes](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Ensembles de conditions requises concernant les commandes de complément](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [Documentation de référence de l’API](../reference/javascript-api-for-office.md)
- [Historique des mises à jour d’Office 365 ProPlus](/officeupdates/update-history-office365-proplus-by-date)
- [Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)](/officeupdates/update-history-office-2019)
- [Historique des mises à jour d’Office 2013 (Démarrer en un clic)](/officeupdates/update-history-office-2013)
- [Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)](/officeupdates/office-updates-msi)
- [Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)](/officeupdates/outlook-updates-msi)
- [Historique des mises à jour d’Office pour Mac](/officeupdates/update-history-office-for-mac)
- [Création de compléments Office](../overview/office-add-ins-fundamentals.md)