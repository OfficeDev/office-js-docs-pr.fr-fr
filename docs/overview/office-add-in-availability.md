---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: 'Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint et OneNote.'
ms.date: 12/01/2017
---

# <a name="office-add-in-host-and-platform-availability"></a>Disponibilité des compléments Office sur les plateformes et les hôtes

Pour fonctionner comme prévu, il se peut que votre complément Office dépende d’un hôte Office spécifique, d’un ensemble de conditions requises, d’un membre d’API ou d’une version de l’API. Les tableaux suivants contiennent la plateforme disponible, les points d’extension, les ensembles de conditions requises de l’API et les ensembles de conditions requises des API communes qui sont actuellement pris en charge pour chaque application Office. 

Si une cellule de tableau contient un astérisque (*), cela signifie que nous travaillons sur celle-ci. Pour les ensembles de conditions requises pour Project ou Access, consultez les [ensembles de conditions requises communs à Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).  

> [!NOTE]
> Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.

## <a name="excel"></a>Excel

<table style="width:80%">
  <tr>
    <th style="width:10%">Plate-forme</th>
    <th style="width:10%">Points d’extension</th> 
    <th style="width:20%">Ensembles de conditions requises de l’API</th> 
    <th style="width:40%"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></th> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Volet Office<br>
        - Contenu<br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </td>
    <td>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
        - BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - CompressedFile<br>
        - Settings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 pour Windows</td>
    <td>
        - Volet Office<br>
        - Contenu</td>
    <td>  - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
        - BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - Settings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 pour Windows</td>
    <td>- Volet Office<br>
        - Contenu<br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td>- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - Settings<br>
        - TextCoercion</td> 
  </tr>
  <tr>
    <td>Office pour iOS</td>
    <td>- Volet Office<br>
        - Contenu</td>
    <td>- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - Settings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 pour Mac</td>
    <td>- Volet Office<br>
        - Contenu<br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td>- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
</table>

<br/>

## <a name="outlook"></a>Outlook

<table style="width:80%">
  <tr>
    <th>Plate-forme</th>
    <th>Points d’extension</th> 
    <th>Ensembles de conditions requises de l’API</th> 
    <th><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></th> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Lecture de message<br>
      - Composition de message<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - Mailbox 1.0<br>
      - Mailbox 1.1<br>
      - Mailbox 1.2<br>
      - Mailbox 1.3<br>
      - Mailbox 1.4<br>
      - Mailbox 1.5</td>
    <td>non disponible</td>
  </tr>
  <tr>
    <td>Office 2013 pour Windows</td>
    <td> - Lecture de message<br>
      - Composition de message<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - Mailbox 1.0<br>
      - Mailbox 1.1<br>
      - Mailbox 1.2<br>
      - Mailbox 1.3<br>
      - Mailbox 1.4</td>
    <td>non disponible</td>
  </tr>
  <tr>
    <td>Office 2016 pour Windows</td>
    <td> - Lecture de message<br>
      - Composition de message<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a><br>
      - Modules</td>
    <td> - Mailbox 1.0<br>
      - Mailbox 1.1<br>
      - Mailbox 1.2<br>
      - Mailbox 1.3<br>
      - Mailbox 1.4<br>
      - Mailbox 1.5</td></td>
    <td>non disponible</td> 
  </tr>
  <tr>
    <td>Office pour iOS</td>
    <td> - Lecture de message<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - Mailbox 1.4</td>
    <td>non disponible</td>
  </tr>
  <tr>
    <td>Office 2016 pour Mac</td>
    <td> - Lecture de message<br>
      - Composition de message<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - Mailbox 1.0<br>
      - Mailbox 1.1<br>
      - Mailbox 1.2<br>
      - Mailbox 1.3<br>
      - Mailbox 1.4<br>
      - Mailbox 1.5</td>
    <td>non disponible</td>
  </tr>
  <tr>
    <td>Office pour Android</td>
    <td> - Lecture de message<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - Mailbox 1.4</td>
    <td>non disponible</td>
  </tr>
</table>

<br/>

## <a name="word"></a>Word

<table style="width:80%">
  <tr>
    <th>Plate-forme</th>
    <th>Points d’extension</th> 
    <th>Ensembles de conditions requises de l’API</th> 
    <th><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></th> 
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Volet Office<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CustomXmlParts<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - DocumentEvents<br>
         - TextFile<br>
         - ImageCoercion<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 pour Windows</td>
    <td> - Volet Office</td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlPart<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - OoxmlCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextFile<br>
         - Settings<br>
         - TextCoercion<br>
         - MatrixCoercion<br>
         - MatrixBindings</td>
  </tr>
  <tr>
    <td>Office 2016 pour Windows</td>
    <td> - Volet Office<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlPart<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - OoxmlCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextFile<br>
         - Settings<br>
         - TextCoercion<br>
         - MatrixCoercion<br>
         - MatrixBindings </td> 
  </tr>
  <tr>
    <td>Office pour iOS</td>
    <td> - Volet Office</td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlPart<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - OoxmlCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextFile<br>
         - Settings<br>
         - TextCoercion<br>
         - MatrixCoercion<br>
         - MatrixBindings </td> 
  </tr>
  <tr>
    <td>Office 2016 pour Mac</td>
    <td> - Volet Office<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlPart<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - OoxmlCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextFile<br>
         - Settings<br>
         - TextCoercion<br>
         - MatrixCoercion<br>
         - MatrixBindings </td> 
  </tr>
</table>

<br/>

## <a name="powerpoint"></a>PowerPoint

<table style="width:80%">
  <tr>
    <th>Plate-forme</th>
    <th>Points d’extension</th> 
    <th>Ensembles de conditions requises de l’API</th> 
    <th><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></th> 
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Contenu<br>
         - Volet Office<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - File<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 pour Windows</td>
    <td> - Contenu<br>
         - Volet Office<br>
    </td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 pour Windows</td>
    <td> - Contenu<br>
         - Volet Office<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office pour iOS</td>
    <td> - Contenu<br>
         - Volet Office</td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
     <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 pour Mac</td>
    <td> - Contenu<br>
         - Volet Office<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
</table>

<br/>

## <a name="onenote"></a>OneNote

<table style="width:80%">
  <tr>
    <th>Plate-forme</th>
    <th>Points d’extension</th> 
    <th>Ensembles de conditions requises de l’API</th> 
    <th><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></th> 
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Contenu<br>
         - Volet Office<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - DocumentEvents<br>
         - Settings<br>
         - TextCoercion<br>
         - HtmlCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 pour Windows</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr> 
  <tr>
    <td>Office 2016 pour Windows</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td> 
  </tr>
  <tr>
    <td>Office pour iOS</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
  <tr>
    <td>Office 2016 pour Mac</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
</table>

<br/>

\* = Nous y travaillons. 

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la plateforme des compléments Office](office-add-ins.md)
- [Ensembles de conditions requises des API communes](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [Ensembles de conditions requises concernant les commandes de complément](https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets)
- [Référence de l’API JavaScript pour Office](https://dev.office.com/reference/add-ins/javascript-api-for-office)

