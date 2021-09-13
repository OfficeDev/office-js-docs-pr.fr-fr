---
title: Créer de meilleurs compléments pour Word avec Office Open XML
description: Vue d’ensemble de l’amélioration de votre add-in Word avec Office Open XML.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 21a70b2b76ef306c06b0b85db5e579fbc1b70eba
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153056"
---
# <a name="create-better-add-ins-for-word-with-office-open-xml"></a>Créer de meilleurs compléments pour Word avec Office Open XML

**Fourni par :**    Stephanie Krieger Microsoft Corporation | Juan Balmori Labra, Microsoft Corporation

Si vous construisez des Office à exécuter dans Word, vous savez peut-être déjà que l’API JavaScript (Office.js) Office propose plusieurs formats pour la lecture et l’écriture de contenu de document. Ces types sont appelés types de contrainte et incluent du texte brut, des tableaux, du code HTML et Office Open XML.

Quelles sont donc les options disponibles pour ajouter du contenu riche à un document, tel que des images, des tableaux mis en forme, des graphiques ou simplement du texte mis en forme ?
Utilisez le code HTML pour insérer certains types de contenu enrichi, tels que des images. En fonction de votre scénario, le forçage HTML peut présenter des inconvénients, tels que les limites de la mise en forme et des options de positionnement disponibles pour votre contenu.
Office Open XML étant le langage dans lequel les documents Word (notamment .docx et .dotx) sont écrits, vous pouvez insérer quasiment tous les types de contenu qu’un utilisateur peut ajouter à un document Word, avec n’importe quel type de mise en forme applicable. Déterminer le balisage Office Open XML nécessaire pour y parvenir est bien plus facile que vous ne le pensez.

> [!NOTE]
> Office Open XML est également le langage utilisé pour les documents PowerPoint et Excel (et pour Visio depuis Office 2013). Cependant, vous pouvez actuellement forcer le contenu au format Office Open XML uniquement dans les compléments Office créés pour Word. Pour plus d’informations sur Office Open XML, notamment pour consulter la documentation de référence du langage complète, reportez-vous à la rubrique [Ressources supplémentaires](#see-also).

Pour commencer, jetez un œil à quelques-uns des types de contenu que vous pouvez insérer à l’aide du forçage Office Open XML. Téléchargez l’exemple de code [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML), qui contient le balisage Office Open XML et le code Office.js nécessaires pour insérer l’un des exemples suivants dans Word.

> [!NOTE]
> Tout au long de cet article, les **termes types** de contenu et contenu enrichi **font** référence aux types de contenu enrichi que vous pouvez insérer dans un document Word.

*Figure 1. Texte avec mise en forme directe*

![Texte avec mise en forme directe appliquée.](../images/office15-app-create-wd-app-using-ooxml-fig01.png)

Utilisez la mise en forme directe pour spécifier exactement à quoi ressemblera le texte, quelle que soit la mise en forme existante dans le document de l’utilisateur.

*Figure 2. Texte mis en forme avec un style*

![Texte mis en forme avec le style de paragraphe.](../images/office15-app-create-wd-app-using-ooxml-fig02.png)

Utilisez un style pour coordonner automatiquement l’apparence du texte que vous insérez avec le document de l’utilisateur.

*Figure 3. Image simple*

![Image d’un logo.](../images/office15-app-create-wd-app-using-ooxml-fig03.png)

Utilisez la même méthode pour insérer n’importe quel Office format d’image pris en charge.

*Figure 4. Image mise en forme avec des styles d’image et des effets*

![Image mise en forme dans Word.](../images/office15-app-create-wd-app-using-ooxml-fig04.png)

L’ajout d’une mise en forme et d’effets de haute qualité à vos images nécessite beaucoup moins de balises que vous ne le pensez.

*Figure 5. Contrôle de contenu*

![Texte dans un contrôle de contenu lié.](../images/office15-app-create-wd-app-using-ooxml-fig05.png)

Utilisez des contrôles de contenu avec votre add-in pour ajouter du contenu à un emplacement spécifié (lié) plutôt qu’à la sélection.

*Figure 6. Zone de texte avec mise en forme WordArt*

![Texte mis en forme avec des effets de texte WordArt.](../images/office15-app-create-wd-app-using-ooxml-fig06.png)

Les effets de texte sont disponibles dans Word pour le texte situé à l’intérieur d’une zone de texte (comme ici) ou pour un corps de texte classique.

*Figure 7. Forme*

![Forme de dessin dans Word.](../images/office15-app-create-wd-app-using-ooxml-fig07.png)

Insérez des formes de dessin intégrées ou personnalisées, avec ou sans texte et effets de mise en forme.

*Figure 8. Tableau avec une mise en forme directe*

![Tableau mis en forme dans Word.](../images/office15-app-create-wd-app-using-ooxml-fig08.png)

Incluez la mise en forme du texte, des bordures, des ombrages, le resserrage des cellules ou toute mise en forme de tableau dont vous avez besoin.

*Figure 9. Tableau mis en forme avec un style de tableau*

![Tableau formaté avec un style de tableau dans Word.](../images/office15-app-create-wd-app-using-ooxml-fig09.png)

Utilisez des styles de tableau intégrés ou personnalisés aussi facilement qu’un style de paragraphe pour du texte.

*Figure 10. Graphique SmartArt*

![Graphique SmartArt dynamique dans Word.](../images/office15-app-create-wd-app-using-ooxml-fig10.png)

Office offre un large éventail de dispositions de diagrammes SmartArt (et vous pouvez utiliser Office Open XML pour créer les vôtres).

*Figure 11. Graphique*

![Graphique dans Word.](../images/office15-app-create-wd-app-using-ooxml-fig11.png)

Vous pouvez insérer des graphiques Excel sous forme de graphiques dynamiques dans des documents Word, ce qui signifie également que vous pouvez les utiliser dans votre complément pour Word. Comme vous pouvez le constater avec les exemples précédents, vous pouvez utiliser le forçage Office Open XML pour insérer pratiquement n’importe quel type de contenu dans un document. Il existe deux façons simples d’obtenir le balisage Office Open XML dont vous avez besoin. Vous pouvez ajouter votre contenu riche à un document Word vierge, puis enregistrer ce fichier au format Document XML Word, ou utiliser un complément de test avec la méthode [getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) pour récupérer le balisage. Les deux approches fournissent globalement le même résultat.

> [!NOTE]
> Un Office document Open XML est en fait un package compressé de fichiers qui représentent le contenu du document. L’enregistrement du fichier au format Document XML Word vous donne l’intégralité du package Open XMLOffice aplati en un seul fichier XML, qui est également ce que vous obtenez lors de l’utilisation pour récupérer le Office `getSelectedDataAsync` Open XML.

Si vous enregistrez le fichier au format XML à partir de Word, notez qu’il existe deux options sous la liste Enregistrer sous type dans la boîte de dialogue Enregistrer sous pour les fichiers au format .xml format. Veillez à choisir **Document XML Word** et non l’option Word 2003.
Téléchargez l’exemple de code [nommé Word-Add-in-Get-Set-EditOpen-XML,](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)que vous pouvez utiliser comme outil pour récupérer et tester votre code.
Et c’est tout ? Pas tout à fait. Pour un grand nombre de scénarios, vous pouvez utiliser le résultat Office Open XML intégral et aplati obtenu avec l’une des méthodes précédentes et tout fonctionnera. La bonne nouvelle est que vous n’avez probablement pas besoin de la majeure partie de ce markup.
Si vous êtes l’un des nombreux développeurs de applications qui voient le markup Open XML Office pour la première fois, le fait d’essayer de comprendre la quantité considérable de marques que vous obtenez pour l’élément de contenu le plus simple peut sembler écrasant, mais cela ne l’est pas toujours.
Dans cette rubrique, vous allez utiliser certains scénarios courants que nous avons entendus de la communauté des développeurs de Office Pour vous montrer les techniques permettant de simplifier Office Open XML pour une utilisation dans votre application. Nous allons explorer le markup pour certains types de contenu affichés précédemment, ainsi que les informations dont vous avez besoin pour réduire la charge utile open XML Office de données. Nous allons également examiner le code dont vous avez besoin pour insérer du contenu enrichi dans un document à l’emplacement de sélection actif et comment utiliser Office Open XML avec l’objet bindings pour ajouter ou remplacer du contenu à des emplacements spécifiés.

## <a name="explore-the-office-open-xml-document-package"></a>Explorer le package Office document Open XML

Lorsque vous utilisez [getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) pour récupérer une sélection de contenu Office Open XML (ou lorsque vous enregistrez le document au format Document XML Word), ce que vous obtenez n’est pas seulement le balisage qui décrit le contenu sélectionné, mais un document entier avec de nombreux paramètres et options dont vous n’aurez certainement pas besoin. En fait, si vous utilisez cette méthode à partir d’un document qui contient un complément de volet de tâches, le balisage que vous obtenez comprend également votre volet de tâches.

Même un simple package de document Word comprend des composants pour les propriétés du document, les styles, le thème (paramètres de mise en forme), les paramètres web, les polices, en plus d’autres composants pour le contenu réel.

Par exemple, supposons que vous voulez insérer uniquement un paragraphe de texte avec une mise en forme directe, comme indiqué précédemment sur la figure 1. Lorsque vous saisissez le Office Open XML pour le texte formaté à l’aide de , vous voyez une grande quantité de `getSelectedDataAsync` marques de contrôle. Ce balisage comprend un élément de package qui représente un document entier, formé de plusieurs parties (communément appelées composants de document ou, dans Office Open XML, composants de package), listées dans la figure 13. Chaque composant représente un fichier distinct du package.

> [!TIP]
> Modifiez Office de texte Open XML dans un éditeur de texte comme Bloc-notes. Si vous l’ouvrez dans Visual Studio, utilisez **Edit >Advanced > Format Document** (Ctrl+K, Ctrl+D) pour mettre en forme le package pour faciliter la modification. Ensuite, vous pouvez réduire ou développer des parties de document ou des sections de celles-ci, comme indiqué dans la figure 12, pour vérifier et modifier plus facilement le contenu du package Office Open XML. Chaque composant du document commence par une balise **pkg:part**.

*Figure 12. Réduction et développement des composants de package pour faciliter la modification dans Visual Studio*

![Office Extrait de code Open XML pour un élément de package dans Visual Studio.](../images/office15-app-create-wd-app-using-ooxml-fig12.png)

*Figure 13. Composants inclus dans un package de document Word Office Open XML de base*

![Extrait de code Office Open XML pour un composant de package.](../images/office15-app-create-wd-app-using-ooxml-fig13.png)

Avec toutes ces balises, vous serez surpris de découvrir que les seuls éléments dont vous avez réellement besoin pour insérer l’exemple de texte mis en forme sont des parties des composants .rels et document.xml.

> [!NOTE]
> Les deux lignes de balisage situées au-dessus de la balise package (déclarations XML pour la version et l’ID de programme Office) sont supposées lorsque vous utilisez le type de forçage Office Open XML. Vous n’avez donc pas à les inclure. Conservez-les si vous voulez ouvrir le balisage modifié en tant que document Word afin de le tester.

Plusieurs des autres types de contenu présentés au début de cette rubrique nécessitent également des composants supplémentaires (au-delà de ceux de la figure 13), et vous aborderez ceux-ci plus loin dans cette rubrique. En attendant, étant donné que vous verrez la plupart des composants affichés dans la figure 13 dans le markup pour n’importe quel package de document Word, voici un résumé rapide de l’objectif de chacun de ces composants et du moment où vous en avez besoin :

- À l’intérieur de la balise package, le premier composant est le fichier .rels, qui définit les relations entre les composants de niveau supérieur du package (généralement les propriétés du document, la miniature (le cas échéant) et le corps du document principal). Une partie du contenu de ce composant est toujours nécessaire dans votre balisage car vous devez définir la relation entre le composant de document principal (où réside votre contenu) et le package de document.

- Le composant document.xml.rels définit les relations pour les composants supplémentaires requis par le composant document.xml (corps principal), le cas échéant.

   > [!IMPORTANT]
   > Les fichiers .rels de votre package (comme les fichiers .rels de niveau supérieur, document.xml.rels et autres qui s’affichent pour certains types de contenu) représentent un outil extrêmement important que vous pouvez utiliser comme guide pour vous aider à modifier rapidement votre package Office Open XML. Pour en savoir plus sur la façon de procéder, voir [Creating your own markup: best practices](#create-your-own-markup-best-practices) plus loin dans cette rubrique.

- Le composant document.xml est le contenu du corps principal du document. Les éléments de ce composant sont évidemment nécessaires, car votre contenu apparaît dans ces éléments. Mais vous n’avez pas besoin de tout ce que vous voyez dans ce composant. Nous étudierons cela plus en détail ultérieurement.

- De nombreux composants sont automatiquement ignorés par les méthodes Set lors de l’insertion de contenu dans un document à l’aide du forçage Office Open XML. Vous pouvez également les supprimer. Il s’agit notamment du fichier theme1.xml (thème de mise en forme du document), les composants des propriétés du document (principales, de complément et de miniature) et les fichiers de paramètres (settings, WebSettings et fontTable).

- Dans l’exemple de la figure 1, la mise en forme du texte est appliquée directement (c’est-à-dire que chaque paramètre de police et de mise en forme de paragraphe est appliqué individuellement). Cependant, si vous utilisez un style (par exemple, si vous voulez que votre texte suive automatiquement la mise en forme du style Titre 1 dans le document de destination) comme indiqué précédemment dans la figure 2, vous aurez besoin d’une partie du composant styles.xml, ainsi que de la définition de relation correspondante. Pour plus d’informations, voir la section « Ajouter des objets qui utilisent des [composants Office Open XML](#add-objects-that-use-additional-office-open-xml-parts)».

## <a name="insert-document-content-at-the-selection"></a>Insérer le contenu du document au niveau de la sélection

Jetons un œil aux exigences minimales de balisage Office Open XML pour l’exemple de texte mis en forme de la figure 1 et au code JavaScript nécessaire pour l’insérer à l’emplacement de sélection actif du document.

### <a name="simplified-office-open-xml-markup"></a>Balisage Office Open XML simplifié

Vous avez modifié l’exemple open XML Office présenté ici, comme décrit dans la section précédente, pour laisser uniquement les composants de document requis et les éléments requis dans chacun de ces composants. Vous allez découvrir comment modifier vous-même le marques de révision (et nous vous expliquerons un peu plus les éléments qui restent ici) dans la section suivante de la rubrique.

```XML
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
        <w:body>
          <w:p>
            <w:pPr>
              <w:spacing w:before="360" w:after="0" w:line="480" w:lineRule="auto"/>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
              <w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>
```

> [!NOTE]
> Si vous ajoutez le balisage représenté ici à un fichier XML avec les balises de déclaration XML pour version et mso-application au début du fichier (figure 13), vous pouvez l’ouvrir dans Word comme un document Word. Ou, sans ces balises, vous pouvez toujours l’ouvrir à l’aide **de Fichier**  >  **ouvert** dans Word. Vous verrez le mode de **compatibilité** sur la barre de titre dans Word, car vous avez supprimé les paramètres qui indiquent à Word qu’il s’agit d’un document Word. Étant donné que vous ajoutez ce markup à un document Word existant, cela n’affectera pas du tout votre contenu.

### <a name="javascript-for-using-setselecteddataasync"></a>JavaScript pour l’utilisation de setSelectedDataAsync

Une fois que vous avez enregistrez l’Office Open XML en tant que fichier XML accessible à partir de votre solution, utilisez la fonction suivante pour définir le contenu du texte mis en forme dans le document à l’aide du contrainte Office Open XML.

Dans cette fonction, vous remarquerez que toutes les lignes sauf la dernière sont utilisées pour obtenir votre balisage enregistré afin de l’utiliser dans l’appel de méthode [setSelectedDataAsync](/javascript/api/office/office.document#setSelectedDataAsync_data__options__callback_) à la fin de la fonction. `setSelectedDataASync` vous devez uniquement spécifier le contenu à insérer et le type de contrainte.

> [!NOTE]
> Remplacez _yourXMLfilename_ par le nom et le chemin du fichier XML que vous avez enregistré dans votre solution. Si vous n’êtes pas sûr de l’endroit où inclure les fichiers XML dans votre solution ou de la façon de les référencer dans votre code, reportez-vous à l’exemple de code [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) pour obtenir des exemples correspondants et un exemple pratique du balisage et du code JavaScript indiqués dans cette rubrique.

```js
function writeContent() {
    var myOOXMLRequest = new XMLHttpRequest();
    var myXML;
    myOOXMLRequest.open('GET', 'yourXMLfilename', false);
    myOOXMLRequest.send();
    if (myOOXMLRequest.status === 200) {
        myXML = myOOXMLRequest.responseText;
    }
    Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'ooxml' });
}
```
## <a name="create-your-own-markup-best-practices"></a>Créer votre propre marque : meilleures pratiques

Étudions de plus près le balisage que vous devez insérer dans l’exemple de texte mis en forme précédent.

Pour cet exemple, commencez par supprimer simplement tous les composants de document du package autres que .rels et document.xml. Ensuite, vous allez modifier ces deux parties requises pour simplifier davantage les choses.

> [!IMPORTANT]
> Utilisez les composants .rels comme une carte afin d’évaluer rapidement les éléments inclus dans le package et de déterminer les composants que vous pouvez supprimer complètement (c’est-à-dire, tous les composants non liés à votre contenu ou référencés par votre contenu). N’oubliez pas que chaque composant de document doit avoir une relation définie dans le package et que ces relations apparaissent dans les fichiers .rels. Elles doivent donc toutes être listées dans un fichier .rels, document.xml.rels ou un fichier .rels propre au contenu.

Le balisage suivant montre le composant .rels requis avant la modification. Dans la mesure où nous supprimons les composants de propriété de document principal et de module de la partie de la miniature, vous devez également supprimer ces relations de .rels. Vous remarquerez que cette opération maintient uniquement la relation (avec l’ID de relation « rID1 » dans l’exemple suivant) pour document.xml.

```XML
<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
  <pkg:xmlData>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.emf"/>
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
    </Relationships>
  </pkg:xmlData>
</pkg:part>
```

> [!IMPORTANT]
> Supprimez les relations (c’est-à-dire, la balise **Relationship**) pour tous les composants que vous supprimez complètement du package. L’ajout d’un composant sans relation correspondante ou l’exclusion d’un composant tout en maintenant sa relation dans le package génère une erreur.

Le balisage suivant présente le composant document.xml, qui contient notre exemple de contenu de texte mis en forme, avant la modification.

```XML
<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document mc:Ignorable="w14 w15 wp14" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <w:body>
          <w:p>
            <w:pPr>
              <w:spacing w:before="360" w:after="0" w:line="480" w:lineRule="auto"/>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
              <w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t>
            </w:r>
            <w:bookmarkStart w:id="0" w:name="_GoBack"/>
            <w:bookmarkEnd w:id="0"/>
          </w:p>
          <w:p/>
          <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
            <w:cols w:space="720"/>
          </w:sectPr>
        </w:body>
      </w:document>
    </pkg:xmlData>
</pkg:part>
```

Étant donné document.xml est le principal document dans lequel vous placez votre contenu, prenez une rapide visite de ce dernier. (La figure 14, qui suit cette liste, représente visuellement le rapport entre une partie du contenu de base et les balises de mise en forme, qui font l’objet de cette rubrique, et ce qui apparaît dans un document Word.)

- La balise de début  **w:document** comprend plusieurs listes d’espaces de noms (**xmlns**). Un grand nombre de ces espaces de noms se réfèrent à des types de contenu spécifiques, dont vous avez besoin uniquement s’ils correspondent à votre contenu.

    Notez que le préfixe des balises dans l’ensemble d’une partie de document fait référence aux espaces de noms. Dans cet exemple, le seul préfixe utilisé dans les balises dans la partie document.xml est **w:**, donc le seul espace de noms que vous devez laisser dans la balise **w:document** d’ouverture est **xmlns:w**.

> [!TIP]
> Si vous modifiez votre balisage dans Visual Studio, après la suppression d’espaces de noms dans un composant, examinez toutes les balises de ce composant. Si vous avez supprimé un espace de noms requis pour votre balisage, un soulignement rouge ondulé apparaît au niveau du préfixe en question pour les balises affectées. Si vous supprimez l’espace de noms **xmlns:mc**, vous devez aussi supprimer l’attribut **mc:Ignorable** qui précède les listes d’espaces de noms.

- Dans la balise Body de début, se trouve une balise de paragraphe (**w:p**) qui comprend le contenu de cet exemple.

- La balise **w:pPr** contient les propriétés de mise en forme directe de paragraphe, comme l’espace avant ou après le paragraphe, l’alignement du paragraphe ou les retraits. (La mise en forme directe se réfère aux attributs que vous appliquez individuellement au contenu plutôt qu’à une partie d’un style). Cette balise comprend également la mise en forme directe de police qui est appliquée à l’ensemble d’un paragraphe, dans une balise imbriquée **w:rPr** (propriétés d’exécution) qui contient la couleur et la taille de la police pour notre exemple.

   > [!NOTE]
   > Vous remarquerez que les tailles de police et les autres paramètres de mise en forme du balisage Word Office Open XML semblent deux fois plus grands que la taille réelle. Ceci est dû au fait que le paragraphe et l’espacement des lignes, ainsi que certaines propriétés de mise en forme de section figurant dans le balisage précédent, sont indiqués en twips (un vingtième de point). Selon les types de contenu avec lesquels vous travaillez dans Office Open XML, vous pouvez voir plusieurs unités de mesure supplémentaires, y compris les unités métriques anglaises (914 400 EMU pour 1 pouce), qui sont utilisées pour certaines valeurs Office Art (drawingML), et la valeur réelle multipliée par 100 000, qui est utilisée pour le balisage drawingML et PowerPoint. PowerPoint exprime également certaines valeurs à 100 fois la valeur réelle, et Excel utilise fréquemment des valeurs réelles.

- Dans un paragraphe, tout contenu avec des propriétés similaires est inclus dans une exécution (**w:r**), comme c’est le cas pour le texte d’exemple. À chaque modification du type de mise en forme ou de contenu, une nouvelle exécution démarre. (C’est-à-dire que si un seul mot du texte d’exemple est en gras, il sera mis de côté pour avoir sa propre exécution.) Dans cet exemple, le contenu inclut uniquement l’exécution de texte.

    Notez que, du fait que la mise en forme incluse dans cet exemple est une mise en forme de police (mise en forme qui peut être appliquée à un seul caractère), elle apparaît également dans les propriétés d’exécution individuelle.

- Examinez également les balises du signet masqué « _GoBack » (**w:bookmarkStart** et **w:bookmarkEnd**), qui apparaissent dans les documents Word par défaut. Vous pouvez toujours supprimer les balises de début et de fin pour le signet GoBack dans votre balisage.

- La dernière partie du corps du document est la balise **w:sectPr**, ou propriétés de section. Cette balise inclut des paramètres tels que les marges et l’orientation de la page. Le contenu que vous insérez à l’aide de **setSelectedDataAsync** applique les propriétés de section actives dans le document de destination par défaut. Ainsi, sauf si votre contenu comprend un saut de section (dans ce cas, vous devez voir plusieurs balises **w:sectPr**), vous pouvez supprimer cette balise.

*Figure 14. Lien entre les balises communes dans document.xml et le contenu ainsi que la mise en page d’un document Word*

![Éléments Office Open XML dans un document Word.](../images/office15-app-create-wd-app-using-ooxml-fig14.png)

> [!TIP]
> Dans le balisage que vous créez, vous pouvez voir un autre attribut dans plusieurs balises, qui comprend les caractères **w:rsid** qui n’apparaissent pas dans les exemples utilisés dans cette rubrique. Il s’agit d’identificateurs de révision. Ils sont utilisés dans Word pour la fonctionnalité Combiner des documents et ils sont activés par défaut. Vous n’en aurez jamais besoin dans le balisage que vous insérez avec votre complément. Les désactiver permet de rendre votre balisage plus lisible. Vous pouvez facilement supprimer les balises RSID existantes ou désactiver la fonctionnalité (comme décrit dans la procédure ci-dessous) pour éviter qu’elles ne soient ajoutées à votre balisage pour le nouveau contenu.

N’oubliez pas que si vous utilisez les fonctionnalités de co-création dans Word (comme la possibilité de modifier simultanément des documents avec d’autres personnes), vous devez activer à nouveau la fonctionnalité lorsque la génération de balisage pour votre complément est terminée.

Pour désactiver les attributs RSID dans Word pour les documents que vous créerez à l’avenir, procédez comme suit :

1. Dans Word, sélectionnez **Fichier**, puis sélectionnez **Options**.
2. Dans la boîte de dialogue Options Word, choisissez **Centre de gestion de la confidentialité**, puis **Paramètres du Centre de gestion de la confidentialité**.
3. Dans la boîte de dialogue Centre de gestion de la confidentialité, sélectionnez **Options de confidentialité**, puis désactivez le paramètre **Stocker un nombre aléatoire pour améliorer l’exactitude de la combinaison**.

Pour supprimer des balises RSID d’un document existant, essayez le raccourci suivant avec le document ouvert dans Office Open XML.

1. Avec le point d’insertion dans le corps principal du document, appuyez sur **Ctrl+Origine** pour accéder au haut du document.
2. Sur le clavier, appuyez sur les touches **Barre d’espace**, **Supprimer**, **Barre d’espace**, puis enregistrez le document.

Après avoir supprimé la majorité du balisage de ce package, nous nous retrouvons avec le balisage minimal qui doit être inséré pour l’exemple, comme indiqué dans la section précédente.

## <a name="use-the-same-office-open-xml-structure-for-different-content-types"></a>Utiliser la même structure Office Open XML pour différents types de contenu

Plusieurs types de contenu riche exigent uniquement les composants .rels et document.xml indiqués dans l’exemple précédent, notamment les contrôles de contenu, les formes de dessin et les zones de texte Office, ainsi que les tableaux (sauf si un style est appliqué au tableau). En effet, vous pouvez réutiliser les mêmes composants de package modifiés et transférer uniquement le contenu **body** de document.xml pour le balisage de votre contenu.

Pour vérifier le balisage Office Open XML pour les exemples de chacun des types de contenu présentés précédemment dans les figures 5 à 8, explorez l’exemple de code [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) mentionné dans la section de présentation.

Avant de passer à autre chose, prenez en compte les différences à noter pour deux de ces types de contenu et la façon d’échanger les éléments dont vous avez besoin.

### <a name="understand-drawingml-markup-office-graphics-in-word-what-are-fallbacks"></a>Comprendre le markup drawingML (Office graphiques) dans Word : qu’est-ce que les fallbacks ?

Si le balisage pour votre forme ou zone de texte semble beaucoup plus complexe que ce à quoi vous vous attendiez, il y a une raison. Avec la version Office 2007, nous avons introduit les formats Office Open XML ainsi qu’un nouveau logiciel graphique Office totalement adopté par PowerPoint et Excel. Dans la version 2007, Word n’utilise qu’une partie de ce logiciel graphique, puisqu’il adopte le logiciel graphique Excel mis à jour, les graphiques SmartArt et des outils d’image avancés. Pour les formes et les zones de texte, Word 2007 a continué à utiliser les anciens objets dessin (VML). C’est dans la version 2010 que Word est passé à l’étape supérieure avec le logiciel graphique qui incorpore les formes et les outils de dessin mis à jour.

Donc, pour prendre en charge les formes et les zones de texte dans des documents Word au format Office Open XML sous Word 2007, les formes (y compris les zones de texte) nécessitent un balisage VML de secours.

En général, comme vous pouvez le voir pour les exemples de forme et de zone de texte inclus dans l’exemple de code [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML), le balisage de secours peut être supprimé. Word ajoute automatiquement le balisage de secours manquant aux formes lorsqu’un document est enregistré. Mais si vous préférez conserver le balisage de secours pour garantir la prise en charge de tous les scénarios utilisateur, c’est tout à fait possible.

Si vous avez regroupé des objets dessin inclus dans votre contenu, un balisage supplémentaire (apparemment répétitif) s’affiche. Celui-ci doit être conservé. Des portions du balisage pour les formes de dessin sont dupliquées lorsque l’objet est inclus dans un groupe.

> [!IMPORTANT]
> Lorsque vous travaillez avec des zones de texte et des formes de dessin, n’oubliez pas de vérifier soigneusement les espaces de noms avant de les supprimer de document.xml. (Si vous réutilisez le balisage d’un autre type d’objet, veillez à rajouter les espaces de noms requis, que vous avez peut-être déjà supprimés de document.xml.) Une grande partie des espaces de noms inclus par défaut dans document.xml sont là pour satisfaire aux exigences des objets dessin.

#### <a name="about-graphic-positioning"></a>Remarque à propos du positionnement des graphiques

Dans les exemples de code [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) et [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML), la zone de texte et la forme sont configurées à l’aide de différents types d’habillage du texte et paramètres de positionnement. (Sachez aussi que les exemples d’image dans ces exemples de code sont configurés en ligne avec la mise en forme du texte, qui positionne un objet graphique sur la ligne de base du texte.)

La forme de ces exemples de code est positionnée par rapport aux marges droite et inférieure de la page. Le positionnement relatif permet une coordination plus facile avec la configuration de document inconnu d’un utilisateur, car le système s’adapte aux marges de l’utilisateur. La mise en page risque alors de paraître moins déséquilibrée à cause de paramètres de marges, d’orientation ou de taille du papier non adaptés. Pour conserver les paramètres de positionnement relatif lorsque vous insérez un objet graphique, vous devez conserver la marque de paragraphe (w:p) dans laquelle est stocké le positionnement (désigné dans Word par le terme point d’ancrage). Si vous insérez le contenu dans une marque de paragraphe existante plutôt que d’inclure la vôtre, vous pourriez être en mesure de conserver le même visuel initial, mais de nombreux types de références relatives qui permettent l’ajustement automatique du positionnement par rapport à la mise en page de l’utilisateur peuvent être perdus.

### <a name="work-with-content-controls"></a>Travailler avec des contrôles de contenu

Les contrôles de contenu représentent une fonctionnalité importante dans Word, car ils peuvent grandement améliorer la puissance de votre complément pour Word de multiples façons, y compris en vous donnant la possibilité d’insérer du contenu à des endroits désignés dans le document plutôt qu’à l’emplacement de sélection uniquement.

Dans Word, retrouvez les contrôles de contenu sur l’onglet Développeur du ruban, comme indiqué dans la figure 15.

*Figure 15. Groupe Contrôles de l’onglet Développeur dans Word*

![Groupe de contrôles de contenu sur le ruban Word.](../images/office15-app-create-wd-app-using-ooxml-fig15.png)

Les types de contrôles de contenu dans Word comprennent du texte enrichi, du texte brut, des images, des galeries de blocs de construction, des cases à cocher, des listes déroulantes, des zones de liste modifiable, un sélecteur de dates et des sections extensibles.

- Utilisez la commande **Propriétés**, indiquée sur la figure 15, pour modifier le titre du contrôle et pour définir des préférences, comme le masquage du conteneur de contrôle.

- Activez le **mode Création** pour modifier le contenu d’espace réservé dans le contrôle.

Si votre complément fonctionne avec un modèle Word, vous pouvez inclure des contrôles dans ce modèle pour améliorer le comportement du contenu. Vous pouvez également utiliser une liaison de données XML dans un document Word pour lier les contrôles de contenu aux données, comme les propriétés du document, pour faciliter la réalisation d’un formulaire ou des tâches similaires. (Pour trouver les contrôles déjà liés aux propriétés intégrées du document dans Word, accédez à l’onglet **Insertion**, sous **QuickPart**.)

Lorsque vous utilisez des contrôles de contenu avec votre complément, vous pouvez aussi étendre considérablement les actions que votre complément peut effectuer à l’aide d’un autre type de liaison. Vous pouvez réaliser une liaison avec un contrôle de contenu à partir du complément, puis écrire le contenu dans la liaison plutôt que dans la sélection active.

> [!NOTE]
> Ne confondez pas la liaison de données XML dans Word et la capacité de liaison avec un contrôle via votre complément. Ce sont des fonctionnalités complètement différentes. Cependant, vous pouvez inclure des contrôles de contenu nommés dans le contenu que vous insérez via votre complément à l’aide du forçage OOXML. Utilisez ensuite le code du complément pour effectuer une liaison avec ces contrôles.

Notez également que les liaisons de données XML et Office.js peuvent interagir avec des composants XML personnalisés de votre application. Il est donc possible d’intégrer ces outils puissants. Pour en savoir plus sur le travail avec les composants XML personnalisés dans l’interface API JavaScript Office, voir [Ressources supplémentaires](#see-also) dans cette rubrique.

L’utilisation des liaisons dans votre complément Word est traitée dans la section suivante de cette rubrique. Tout d’abord, prenons un exemple du Office Open XML requis pour l’insertion d’un contrôle de contenu de texte enrichi que vous pouvez lier à l’aide de votre add-in.

> [!IMPORTANT]
> Les contrôles de texte enrichi sont le seul type de contrôle de contenu que vous pouvez utiliser pour effectuer une liaison avec un contrôle de contenu provenant de votre complément.

```XML
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" >
        <w:body>
          <w:p/>
          <w:sdt>
              <w:sdtPr>
                <w:alias w:val="MyContentControlTitle"/>
                <w:id w:val="1382295294"/>
                <w15:appearance w15:val="hidden"/>
                <w:showingPlcHdr/>
              </w:sdtPr>
              <w:sdtContent>
                <w:p>
                  <w:r>
                  <w:t>[This text is inside a content control that has its container hidden. You can bind to a content control to add or interact with content at a specified location in the document.]</w:t>
                </w:r>
                </w:p>
              </w:sdtContent>
            </w:sdt>
          </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
 </pkg:package>
```

Comme mentionné précédemment, les contrôles de contenu, comme le texte mis en forme, ne nécessitent aucun composant de document supplémentaire, de sorte que seules les versions modifiées des composants .rels et document.xml sont incluses.

La balise **w:sdt** dans le corps de document.xml représente le contrôle de contenu. Si vous générez le balisage Office Open XML pour un contrôle de contenu, vous verrez que plusieurs attributs ont été supprimés de cet exemple, y compris la balise et les propriétés de composant de document. Seuls les éléments essentiels (et quelques éléments recommandés) ont été conservés, notamment les éléments suivants :

- **L’alias** est la propriété de titre de la boîte de dialogue Propriétés du contrôle de contenu dans Word. Cette propriété est obligatoire (elle représente le nom de l’élément) si vous envisagez une liaison au contrôle à partir de votre complément.

- 
            **id** (unique) est une propriété obligatoire. Si vous liez le contrôle à partir de votre complément, l’ID est la propriété que la liaison utilise dans le document pour identifier le contrôle de contenu nommé applicable.

- **L’attribut** d’apparence est utilisé pour masquer le conteneur de contrôles, pour une apparence plus propre. Cette fonctionnalité a été introduite dans Word 2013, comme vous le voyez par l’utilisation de l’espace de noms w15. Cette propriété étant utilisée, l’espace de noms w15 est conservé au début du composant document.xml.

- **L’attribut showingPlcHdr** est un paramètre facultatif qui définit le contenu par défaut que vous incluez dans le contrôle (texte dans cet exemple) en tant que contenu d’espace réservé. Ainsi, si l’utilisateur clique ou appuie dans la zone de contrôle, tout le contenu est sélectionné au lieu de se comporter comme du contenu modifiable dans lequel l’utilisateur peut apporter des modifications.

- Bien que la marque de paragraphe vide (**w:p/**) qui précède la balise **sdt** ne soit pas nécessaire pour ajouter un contrôle de contenu (et ajoute un espace vertical au-dessus du contrôle dans le document Word), elle garantit que le contrôle est placé dans son propre paragraphe. Cela peut être important, selon le type et la mise en forme du contenu qui sera ajouté dans le contrôle.

- Si vous envisagez de lier le contrôle, le contenu par défaut du contrôle (à l’intérieur de la balise **sdtContent**) doit comprendre au moins un paragraphe complet (comme dans cet exemple), pour que votre liaison accepte le contenu riche à plusieurs paragraphes.

> [!NOTE]
> L’attribut de composant de document qui a été supprimé de cette balise d’exemple **w:sdt** peut apparaître dans un contrôle de contenu pour référencer un composant distinct dans le package où les informations de contenu d’espace réservé peuvent être stockées (composants situés dans un répertoire de glossaire du package Office Open XML). Bien que « composant de document » soit le terme utilisé pour les composants XML (autrement dit, les fichiers) d’un package Office Open XML, le terme « composants de document » tel qu’il est utilisé dans la propriété sdt fait référence au même terme dans Word, utilisé pour décrire des types de contenu, notamment les blocs de construction et les composants QuickPart des propriétés de document (par exemple, les contrôles XML liés aux données et intégrés). Si des composants existent sous un répertoire de glossaire dans votre package Office Open XML, il se peut que vous deviez les conserver si le contenu que vous insérez comprend ces fonctionnalités. Pour un contrôle de contenu typique que vous souhaitez utiliser pour une liaison à partir de votre complément, celles-ci ne sont pas requises. N’oubliez pas que si vous supprimez les composants de glossaire du package, vous devez également supprimer l’attribut de composant de document de la balise w:sdt.

La section suivante sera consacrée à la création et à l’utilisation de liaisons dans votre complément Word.

## <a name="insert-content-at-a-designated-location"></a>Insérer du contenu à un emplacement désigné

Vous avez déjà vu comment insérer du contenu au niveau de la sélection active dans un document Word. Si vous établissez une liaison avec un contrôle de contenu nommé qui figure dans le document, vous pouvez insérer n’importe lequel de ces types de contenu dans ce contrôle.

Quand utiliser cette approche ?

- Lorsque vous avez besoin d’ajouter ou de remplacer du contenu à des emplacements spécifiés dans un modèle, notamment pour remplir des parties du document à partir d’une base de données

- Lorsque vous voulez avoir la possibilité de remplacer le contenu que vous insérez à l’emplacement de sélection actif, notamment pour fournir des options d’élément de conception à l’utilisateur

- Lorsque vous voulez que l’utilisateur ajoute au document des données auxquelles vous pouvez accéder pour une utilisation avec votre complément, notamment pour remplir les champs dans le volet de tâches en fonction des informations que l’utilisateur ajoute dans le document

Téléchargez l’exemple de code [Word-Add-in-JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings), qui fournit un exemple pratique décrivant comment insérer et lier un contrôle de contenu, et comment remplir la liaison.

### <a name="add-and-bind-to-a-named-content-control"></a>Ajout et liaison à un contrôle de contenu nommé

En examinant le code JavaScript qui suit, prenez en compte ces exigences :

- Comme mentionné précédemment, vous devez utiliser un contrôle de contenu de texte enrichi afin d’établir une liaison avec le contrôle à partir de votre complément Word.

- Le contrôle de contenu doit avoir  un nom (il s’agit du champ Titre dans la boîte de dialogue Propriétés du contrôle de contenu, qui correspond à la balise **Alias** dans le Office open XML). Voici comment le code détermine où placer la liaison.

- Vous pouvez disposer de plusieurs contrôles nommés et les lier en fonction de vos besoins. Utilisez un nom de contrôle de contenu unique, des ID de contrôle de contenu uniques et un ID de liaison unique.

```js
function addAndBindControl() {
    Office.context.document.bindings.addFromNamedItemAsync("MyContentControlTitle", "text", { id: 'myBinding' }, function (result) {
        if (result.status == "failed") {
            if (result.error.message == "The named item does not exist.")
                var myOOXMLRequest = new XMLHttpRequest();
                var myXML;
                myOOXMLRequest.open('GET', '../../Snippets_BindAndPopulate/ContentControl.xml', false);
                myOOXMLRequest.send();
                if (myOOXMLRequest.status === 200) {
                    myXML = myOOXMLRequest.responseText;
                }
                Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'ooxml' }, function (result) {
                    Office.context.document.bindings.addFromNamedItemAsync("MyContentControlTitle", "text", { id: 'myBinding' });
                });
        }
    });
}
```

Le code présenté ici suit les étapes ci-après.

- Tentative de création d’une liaison avec le contrôle de contenu nommé, à l’aide de [addFromNamedItemAsync](/javascript/api/office/office.bindings#addFromNamedItemAsync_itemName__bindingType__options__callback_).

  Effectuez d’abord cette opération s’il est possible que le contrôle nommé existe déjà dans le document lors de l’exécution du code. Par exemple, vous devez procéder de cette façon si le complément a été inséré et enregistré dans un modèle conçu pour fonctionner avec le complément dans lequel le contrôle a été placé à l’avance. Vous devez également procéder ainsi si vous devez créer une liaison à un contrôle qui a été placé précédemment par le complément.

- Le rappel dans le premier appel à la méthode vérifie l’état du résultat pour voir si la liaison a échoué car l’élément nommé n’existe pas dans le document (autrement dit, le contrôle de contenu nommé `addFromNamedItemAsync` MyContentControlTitle dans cet exemple). Si c’est le cas, le code ajoute le contrôle au point de sélection actif (à l’aide de ), puis `setSelectedDataAsync` s’y lie.

> [!NOTE]
> Comme mentionné plus haut et illustré dans le code précédent, le nom du contrôle de contenu est utilisé pour déterminer où créer la liaison. Cependant, dans le balisage Office Open XML, le code ajoute la liaison au document en utilisant à la fois le nom et l’attribut ID du contrôle de contenu.

Après l’exécution du code, si vous examinez le balisage du document dans lequel votre complément a créé des liaisons, vous voyez deux composants pour chaque liaison. Dans le markup du contrôle de contenu où une liaison a été ajoutée (dans document.xml), vous verrez l’attribut **w15:webExtensionLinked/**.

Dans la partie de document nommée webExtensions1.xml, vous voyez la liste des liaisons que vous avez créées. Chacun d’eux est identifié à l’aide de l’ID de liaison et de l’attribut ID du contrôle applicable, par exemple, où l’attribut **appref** est l’ID du contrôle de contenu : **we:binding id="myBinding » type="text » appref="1382295294"/**.

> [!IMPORTANT]
> Vous devez ajouter la liaison au moment où vous avez l’intention d’agir dessus. N’incluez pas le balisage pour la liaison dans le code Office Open XML pour l’insertion du contrôle de contenu car le processus d’insertion de ce balisage supprimera la liaison.

### <a name="populate-a-binding"></a>Remplissage d’une liaison

Le code pour l’écriture du contenu d’une liaison est semblable à celui de l’écriture du contenu d’une sélection.

```js
function populateBinding(filename) {
  var myOOXMLRequest = new XMLHttpRequest();
  var myXML;
  myOOXMLRequest.open('GET', filename, false);
  myOOXMLRequest.send();
  if (myOOXMLRequest.status === 200) {
      myXML = myOOXMLRequest.responseText;
  }
  Office.select("bindings#myBinding").setDataAsync(myXML, { coercionType: 'ooxml' });
}
```

Comme avec `setSelectedDataAsync` , vous spécifiez le contenu à insérer et le type de contrainte. La seule exigence supplémentaire pour l’écriture sur une liaison est l’identification de la liaison par un ID. Notez comment l’ID de liaison utilisé dans ce code (bindings#myBinding) correspond à l’ID de liaison établi (myBinding) lors de la création de la liaison dans la fonction précédente.

> [!NOTE]
> Le code précédent est l’unique élément nécessaire pour le remplissage initial ou le remplacement du contenu dans une liaison. Lorsque vous insérez un contenu nouveau à un emplacement lié, le contenu existant de cette liaison est automatiquement remplacé. Découvrez un exemple de cette opération dans l’exemple de code précédemment cité [Word-Add-in-JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings), qui fournit deux exemples de contenu distincts que vous pouvez utiliser indifféremment pour remplir la même liaison.

## <a name="add-objects-that-use-additional-office-open-xml-parts"></a>Ajouter des objets qui utilisent des composants Office Open XML

De nombreux types de contenu nécessitent des composants de document supplémentaires dans le package Office Open XML, ce qui signifie qu’ils référencent des informations dans un autre composant ou que le contenu lui-même est stocké dans un ou plusieurs composants supplémentaires et référencé dans document.xml.

Par exemple, tenez compte des éléments suivants :

- Le contenu qui utilise des styles de mise en forme (comme le texte avec style de la figure 2 ou le tableau avec style de la figure 9) nécessite le composant styles.xml.

- Les images (comme celles des figures 3 et 4) comprennent des données d’image binaires dans un (et parfois deux) composant supplémentaire.

- Les diagrammes SmartArt (comme celui de la figure 10) nécessitent plusieurs composants supplémentaires pour décrire la mise en page et le contenu.

- Les graphiques (comme celui de la figure 11) nécessitent plusieurs composants supplémentaires, y compris leur propre composant de relation (.rels).

Vous pouvez voir des exemples modifiés de balisage pour tous ces types de contenu dans l’exemple de code précédemment cité [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML). Vous pouvez insérer tous ces types de contenu avec le code JavaScript mentionné plus haut (et indiqué dans les exemples de code référencés) pour l’insertion de contenu à l’emplacement de sélection actif et l’écriture de contenu sur un emplacement spécifié à l’aide des liaisons.

Avant d’explorer les exemples, prenons quelques conseils pour travailler avec chacun de ces types de contenu.

> [!IMPORTANT]
> N’oubliez pas que si vous conservez des composants supplémentaires référencés dans document.xml, vous devrez conserver document.xml.rels et les définitions de relation pour les composants applicables que vous conservez, comme styles.xml ou un fichier image.

### <a name="working-with-styles"></a>Utilisation des styles

La même approche de modification du marques de révision que dans l’exemple précédent avec du texte directement formaté s’applique lorsque vous utilisez des styles de paragraphe ou des styles de tableau pour mettre en forme votre contenu. Toutefois, le balisage utilisé pour les styles de paragraphe est bien plus simple, comme le montre l’exemple décrit ici.

#### <a name="editing-the-markup-for-content-using-paragraph-styles"></a>Modification du balisage du contenu avec des styles de paragraphe

Le balisage suivant représente le contenu du corps pour l’exemple de texte avec le style indiqué sur la figure 2.

```XML
<w:body>
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Heading1"/>
    </w:pPr>
    <w:r>
      <w:t>This text is formatted using the Heading 1 paragraph style.</w:t>
    </w:r>
  </w:p>
</w:body>
```

> [!NOTE]
> Comme vous pouvez le voir, le balisage du texte mis en forme dans document.xml est beaucoup plus simple lorsque vous utilisez un style, car ce dernier contient l’ensemble de la mise en forme de paragraphe et de police qu’il vous faudrait sinon référencer individuellement. Cependant, comme expliqué précédemment, si vous voulez utiliser des styles ou une mise en forme directe à des fins différentes : utilisez la mise en forme directe pour spécifier l’apparence de votre texte indépendamment de la mise en forme dans le document de l’utilisateur ; utilisez un style de paragraphe (notamment un nom de style de paragraphe intégré, comme Heading 1 ici) pour obtenir la coordination automatique de la mise en forme du texte avec le document de l’utilisateur.

L’utilisation d’un style illustre très bien l’importance de la lecture et de la compréhension du balisage pour le contenu que vous insérez, car le fait qu’un autre composant de document soit référencé ici n’est pas explicite. Si vous incluez la définition de style dans ce balisage sans inclure le composant styles.xml, les informations de style dans document.xml seront ignorées, indépendamment de l’utilisation ou non de ce style dans le document de l’utilisateur.

Toutefois, si vous regardez le composant styles.xml, vous verrez que seule une petite partie de ce long élément de balisage est nécessaire lors de la modification du balisage pour utilisation dans votre complément :

- Le composant styles.xml inclut plusieurs espaces de noms par défaut. Si vous conservez uniquement les informations de style nécessaires pour votre contenu, dans la plupart des cas, vous ne devrez garder que l’espace de noms **xmlns:w**.

- Le contenu de la balise **w:docDefaults** qui se trouve au début du composant styles sera ignoré lors de l’insertion de votre balisage par le biais du composant. Il peut donc être supprimé.

- Le plus grand élément de balisage dans un composant styles.xml est destiné à la balise **w:latentStyles** qui apparaît après docDefaults, qui fournit des informations (telles que les attributs d’apparence pour le volet Styles et la galerie Styles) pour chaque style disponible. Ces informations sont également ignorées lors de l’insertion de contenu par le biais de votre complément. Elles peuvent donc être supprimées.

- Après les informations de styles latentes, vous voyez une définition pour chaque style utilisé dans le document à partir duquel votre balisage a été généré. Cela inclut certains styles par défaut utilisés lors de la création d’un document et qui peuvent ne pas être appropriés pour votre contenu. Vous pouvez supprimer les définitions de tous les styles qui ne sont pas utilisés par votre contenu.

   > [!NOTE]
   > Chaque style de titre intégré possède un style de caractère associé, qui est une version de style de caractère du même format de titre. Sauf si vous avez appliqué le style de titre en tant que style de caractère, vous pouvez le supprimer. Si le style est utilisé en tant que style de caractère, il apparaît dans document.xml, dans une balise de propriétés d’exécution (**w:rPr**) et non dans la balise de propriétés de paragraphe (**w:pPr**). Cela se produit uniquement si vous avez appliqué le style à une partie d’un paragraphe seulement, mais cela peut arriver si vous avez appliqué le style de façon incorrecte par mégarde.

- Si vous utilisez un style intégré pour votre contenu, vous n’avez pas à inclure une définition complète. Vous devez inclure uniquement le nom du style, l’ID du style et au moins un attribut de mise en forme pour que le code Office Open XML forcé applique le style à votre contenu lors de l’insertion.

    Toutefois, il est recommandé d’inclure une définition de style complète (même s’il s’agit de la valeur par défaut pour les styles intégrés). Si un style est déjà utilisé dans le document de destination, votre contenu prend la définition résidente pour le style, indépendamment de ce que vous incluez dans le fichier styles.xml. Si le style n’est pas encore utilisé dans le document de destination, votre contenu utilisera la définition de style que vous fournissez dans les marques.

Par exemple, le seul contenu à conserver à partir du styles.xml pour l’exemple de texte de la figure 2, qui est mise en forme à l’aide du style Heading 1, est le suivant :

> [!NOTE]
> Une définition Word complète pour le style Heading 1 a été conservée dans cet exemple.

```XML
<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml">
  <pkg:xmlData>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:link w:val="Heading1Char"/>
        <w:uiPriority w:val="9"/>
        <w:qFormat/>
        <w:pPr>
          <w:keepNext/>
          <w:keepLines/>
          <w:spacing w:before="240" w:after="0" w:line="259" w:lineRule="auto"/>
          <w:outlineLvl w:val="0"/>
        </w:pPr>
        <w:rPr>
          <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
          <w:color w:val="2E74B5" w:themeColor="accent1" w:themeShade="BF"/>
          <w:sz w:val="32"/>
          <w:szCs w:val="32"/>
        </w:rPr>
      </w:style>
    </w:styles>
  </pkg:xmlData>
</pkg:part>
```

#### <a name="edit-the-markup-for-content-using-table-styles"></a>Modifier le marques de révision pour le contenu à l’aide de styles de tableau

Lorsque votre contenu utilise un style de tableau, le même composant relatif de styles.xml est nécessaire, comme décrit pour l’utilisation des styles de paragraphe. Autrement dit, il vous suffit de conserver les informations pour le style que vous utilisez dans votre contenu, puis d’inclure le nom, l’ID et au moins un attribut de mise en forme, mais il est préférable d’inclure une définition de style complète pour traiter tous les scénarios utilisateur possibles.

Cependant, lorsque vous examinez le balisage de votre tableau dans document.xml et de votre définition de style de tableau dans styles.xml, vous remarquez que les balises sont beaucoup plus nombreuses que pour les styles de paragraphe.

- Dans document.xml, la mise en forme est appliquée par cellule même si elle est incluse dans un style. L’utilisation d’un style de tableau ne réduit pas le volume du balisage. L’utilisation de styles de tableau pour le contenu a pour avantage de faciliter la mise à jour et la coordination de l’apparence de plusieurs tableaux.

- Dans styles.xml, vous remarquerez également une quantité importante de balisage pour un seul style de tableau, car les styles de tableau comprennent plusieurs types d’attributs de mise en forme possibles pour chacune des différentes zones de table, comme l’ensemble du tableau, les lignes de titre, les colonnes et lignes à bandes paires et impaires (séparément), la première colonne, etc.

### <a name="work-with-images"></a>Travailler avec des images

Le balisage d’une image inclut une référence à au moins un composant qui contient les données binaires permettant de décrire votre image. Pour une image complexe, on peut compter des centaines de pages de balisage non modifiable. Puisque vous n’avez pas à modifier les composants binaires, vous pouvez simplement les réduire si vous utilisez un éditeur structuré tel que Visual Studio, de sorte que vous pouvez toujours facilement vérifier et modifier le reste du package.

Si vous observez l’exemple de balisage de l’image simple montré plus haut sur la figure 3, disponible dans l’exemple de code référencé précédemment [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML), vous verrez que le balisage de l’image dans document.xml comprend des informations sur la taille et la position, ainsi qu’une référence de relation avec le composant qui contient les données d’image binaires. Cette référence est incluse dans la balise **a:blip**, comme suit :

```XML
<a:blip r:embed="rId4" cstate="print">
```

Notez que, du fait qu’une référence de relation est explicitement utilisée ( **r:embed="rID4"** ) et que le composant associé est obligatoire pour l’affichage de l’image, si vous n’incluez pas les données binaires dans votre package Office Open XML, une erreur sera générée. Ce cas de figure ne se présente pas avec styles.xml, comme expliqué précédemment, qui ne générera pas d’erreur en cas d’omission car la relation n’est pas explicitement référencée et est établie avec un composant qui fournit des attributs au contenu (mise en forme) plutôt que d’appartenir au contenu lui-même.

> [!NOTE]
> Lorsque vous examinez le balisage, notez les espaces de noms supplémentaires utilisés dans la balise a:blip. Vous verrez dans document.xml que l’espace de noms **xlmns:a** (l’espace de noms drawingML principal) est placé dynamiquement au début de l’utilisation des références drawingML plutôt qu’en haut de la partie document.xml. Cependant, l’espace de noms de relations (r) doit être conservé lorsqu’il apparaît au début de document.xml. Vérifiez si votre balise d’image comporte des exigences d’espace de noms supplémentaires. N’oubliez pas que vous n’avez pas à mémoriser les types de contenu devant être associés aux espaces de noms : les préfixes des balises vous l’indiquent tout au long du document.xml.

### <a name="understanding-additional-image-parts-and-formatting"></a>Présentation des composants d’image supplémentaires et de la mise en forme

Lorsque vous utilisez des effets de mise en forme d’image Office sur votre image, comme pour l’image de la figure 4, qui utilise la luminosité ajustée et les paramètres de contraste (en plus des styles d’image), un second composant de données binaires pour une copie au format HD des données d’image peut être nécessaire. Ce format HD supplémentaire est requis pour la mise en forme considérée comme un effet de superposition et la référence apparaît dans document.xml, comme suit :

```XML
<a14:imgLayer r:embed="rId5">
```

Voir le balisage nécessaire pour l’image mise en forme sur la figure 4 (qui utilise notamment des effets de superposition) dans l’exemple de code [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML).

### <a name="work-with-smartart-diagrams"></a>Travailler avec des diagrammes SmartArt

Un diagramme SmartArt possède quatre composants associés, mais seulement deux sont toujours requis. Vous pouvez examiner un exemple de balisage SmartArt dans l’exemple de code [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML). Lisez d’abord une brève description de chacun des composants et découvrez pourquoi ils sont requis ou non :

> [!NOTE]
> Si votre contenu comprend plusieurs diagrammes, ils seront numérotés les uns à la suite des autres, en remplaçant le chiffre 1 dans les noms de fichier répertoriés ici.

- layout1.xml : ce composant est requis. Il inclut la définition de balisage pour l’apparence et les fonctionnalités de la mise en page.

- data1.xml : ce composant est requis. Il comprend les données utilisées dans votre instance du diagramme.

- drawing1.xml : ce composant n’est pas toujours requis, mais si vous appliquez une mise en forme personnalisée aux éléments dans l’instance d’un diagramme, tel que la mise en forme directe de formes individuelles, vous devrez peut-être le conserver.

- colors1.xml : ce composant n’est pas requis. Il comprend des informations sur le style de couleur, mais les couleurs de votre diagramme seront coordonnées par défaut avec les couleurs du thème de mise en forme actif dans le document de destination, en fonction du style de couleur SmartArt que vous appliquez à partir des outils SmartArt disponibles dans l’onglet Création dans Word avant l’enregistrement de votre balisage Office Open XML.

- quickStyles1.xml : ce composant n’est pas requis. Comme le composant des couleurs, vous pouvez le supprimer car votre diagramme suivra la définition du style SmartArt appliqué disponible dans le document de destination (autrement dit, la coordination sera automatique avec le thème de mise en forme dans le document de destination).

> [!TIP]
> Le fichier SmartArt layout1.xml est un bon exemple pour illustrer les parties que vous pouvez supprimer de votre balisage, mais il peut être inutile de consacrer davantage de temps à cela (car cette opération supprime une petite quantité de balisage par rapport à la totalité du package). Si vous voulez vous débarrasser de toutes les lignes possibles de balisage, vous pouvez supprimer la balise **dgm:sampData** et son contenu. Ces données d’exemple définissent l’apparence de la miniature d’aperçu pour le diagramme dans les galeries de styles SmartArt. Toutefois, si elles sont omises, les exemples de données par défaut sont utilisés.

N’ignorez pas que le markup d’un diagramme SmartArt dans document.xml contient des références d’ID de relation à la disposition, aux données, aux couleurs et aux composants de styles rapides. Vous pouvez supprimer les références dans document.xml aux composants de couleurs et de styles lorsque vous supprimez ces composants et leurs définitions de relation (et il est certainement préférable de le faire, étant donné que vous supprimez ces relations), mais vous n’obtenez pas d’erreur si vous les laissez, car elles ne sont pas requises pour que votre diagramme soit inséré dans un document. Recherchez ces références dans document.xml la balise **dgm:relIds.** Que vous passiez ou non cette étape, conservez les références d’ID de relation pour la disposition requise et les composants de données.

### <a name="work-with-charts"></a>Travailler avec des graphiques

Comme les diagrammes SmartArt, les graphiques contiennent plusieurs composants supplémentaires. Cependant, la configuration de ces graphiques est légèrement différente de la configuration SmartArt car les graphiques ont leur propre fichier de relation. Voici une description des composants de document requis et amovibles d’un graphique.

> [!NOTE]
> Comme pour les diagrammes SmartArt, si votre contenu comprend plusieurs graphiques, ils seront numérotés les uns à la suite des autres, le chiffre 1 étant remplacé dans les noms de fichier répertoriés ici.

- Dans document.xml.rels, vous verrez une référence au composant requis qui contient les données décrivant le graphique (chart1.xml).

- Vous noterez également un fichier de relation distinct pour chaque graphique de votre package Office Open XML, comme chart1.xml.rels.

    Il existe trois fichiers référencés dans chart1.xml.rels, mais un seul est requis. Ils comprennent les données de classeur Excel binaires (requises) et les composants de couleurs et de styles (colors1.xml et styles1.xml), que vous pouvez supprimer.

Les graphiques que vous pouvez créer et modifier en mode natif dans Word sont des graphiques Excel. Leurs données sont conservées sur une feuille de calcul Excel qui est incorporée sous forme de données binaires dans votre package Office Open XML. Comme les composants de données binaires pour les images, ces données binaires Excel sont obligatoires, mais rien n’est à modifier dans ce composant. Ainsi, vous pouvez simplement réduire le composant dans l’éditeur pour éviter de devoir tout faire défiler manuellement afin d’examiner le reste de votre package Office Open XML.

Toutefois, comme pour SmartArt, vous pouvez supprimer les composants de couleurs et de styles. Si vous avez utilisé les styles de graphique et les styles de couleurs disponibles dans pour mettre en forme votre graphique, celui-ci adoptera automatiquement la mise en forme applicable lors de son insertion dans le document de destination.

Voir le balisage modifié pour l’exemple de graphique de la figure 11 dans l’exemple de code [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML).

## <a name="edit-the-office-open-xml-for-use-in-your-task-pane-add-in"></a>Modifier le Office Open XML pour l’utiliser dans votre add-in du volet Des tâches

Nous avons déjà vu comment identifier et modifier le contenu de votre balisage. Si la tâche semble encore difficile lorsque vous jetez un œil au package Office Open XML généré pour votre document, voici un résumé rapide des étapes recommandées pour vous aider à modifier ce package rapidement.

> [!NOTE]
> N’oubliez pas que vous pouvez utiliser tous les composants .rels du package comme une carte pour rechercher rapidement les composants de document que vous pouvez supprimer.

1. Ouvrez le fichier XML aplati dans Visual Studio et appuyez sur Ctrl+K ou Ctrl+D pour mettre en forme le fichier. Ensuite, utilisez les boutons Réduire/Développer situés sur la gauche pour réduire les composants que vous devez supprimer. Vous pouvez également réduire les longs composants dont vous avez besoin, mais que vous n’avez pas à modifier (par exemple, les données binaires encodées en base64 pour un fichier image), cela vous permet de lire le balisage plus rapidement et plus facilement.

2. Plusieurs composants du package de document peuvent être quasi systématiquement supprimés quand vous préparez le balisage Office Open XML pour l’utiliser dans votre complément. Commencez d’abord par supprimer ces composants (et leurs définitions de relation associées). Ceci va déjà considérablement réduire la taille du package. Parmi ces composants figurent theme1, fontTable, settings, webSettings, la miniature, les fichiers de propriétés principales et de complément, et tous les composants `taskpane` ou `webExtension`.

3. Supprimez tous les composants qui ne sont pas liés à votre contenu, comme les notes de bas de page, les en-têtes ou les pieds de page dont vous n’avez pas besoin. Là encore, n’oubliez pas de supprimer également les relations associées.

4. Vérifiez le composant document.xml.rels pour déterminer si des fichiers référencés dans ce composant sont requis pour votre contenu, comme un fichier image, le composant styles ou des composants de diagramme SmartArt. La suppression des relations des composants de votre contenu n’exige pas que vous ayez également supprimé le composant associé, et ne confirme pas cette suppression. Si votre contenu ne nécessite aucun des composants de document référencés dans document.xml.rels, vous pouvez également supprimer ce fichier.

5. Si votre contenu possède un composant .rels supplémentaire (comme chart#.xml.rels), vérifiez s’il existe d’autres composants référencés que vous pouvez supprimer (comme les styles rapides pour les graphiques) et supprimez la relation de ce fichier, ainsi que le composant associé.

6. Modifiez document.xml pour supprimer les espaces de noms non référencés dans le composant, les propriétés de section si votre contenu ne comprend pas de saut de section et tout balisage qui n’est pas lié au contenu à insérer. Si vous insérez des formes ou des zones de texte, vous pouvez également supprimer le balisage complet de secours.

7. Modifiez tous les composants supplémentaires requis lorsque vous savez que vous pouvez supprimer le balisage important sans affecter votre contenu, comme le composant styles.

Après les sept étapes précédentes, vous avez supprimé probablement entre 90 et 100 % du balisage supprimable, en fonction de votre contenu. Dans la plupart des cas, vous avez probablement atteint la quantité voulue d’éléments supprimés.

Que vous vous arrêtiez à cette étape ou que vous décidiez de continuer à explorer votre contenu pour trouver les dernières lignes de balisage que vous pouvez supprimer, n’oubliez pas que vous pouvez utiliser l’exemple de code précédemment référencé [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML) comme complément de travail pour tester rapidement et facilement votre balisage modifié.

> [!TIP]
> Si vous mettez à jour un extrait Office Open XML dans une solution existante lors du développement, effacez les fichiers Internet temporaires avant d’exécuter à nouveau la solution pour mettre à jour le balisage Office Open XML utilisé par votre code. Le balisage qui est inclus dans votre solution pour les fichiers XML est mis en cache sur votre ordinateur. Vous pouvez évidemment effacer les fichiers Internet temporaires à partir de votre navigateur web par défaut. Pour accéder aux options Internet et supprimer ces paramètres à partir de Visual Studio 2019, dans le menu **Débogage,** choisissez **Options**. Ensuite, sous **Environnement**, choisissez **Navigateur web**, puis **Options Internet Explorer**.

## <a name="create-an-add-in-for-both-template-and-stand-alone-use"></a>Créer un module pour un modèle et une utilisation autonome

Dans cette rubrique, vous avez découvert plusieurs exemples de ce que vous pouvez faire avec Office Open XML dans vos compléments pour . Vous avez examiné un large éventail d’exemples de type de contenu enrichi que vous pouvez insérer dans des documents à l’aide du type de forage Office Open XML, ainsi que des méthodes JavaScript pour insérer ce contenu à l’emplacement de sélection ou à un emplacement spécifié (lié).

Alors, qu’avez-vous besoin de savoir en plus si vous créez votre complément pour une utilisation autonome (c’est-à-dire, insérée à partir de l’ Store ou d’un emplacement de serveur propriétaire) et pour une utilisation dans un modèle pré-créé, conçu pour travailler avec votre complément ? En réalité, vous savez déjà tout.

Le balisage pour un type de contenu donné et les méthodes pour l’insérer sont les mêmes que votre complément soit conçu pour une utilisation autonome ou avec un modèle. Si vous utilisez des modèles conçus pour fonctionner avec votre complément, assurez-vous simplement que votre code JavaScript comprend des rappels prenant en compte les scénarios dans lesquels le contenu référencé est susceptible d’être déjà présent dans le document (comme démontré dans l’exemple de liaison présenté dans la section [Ajout et liaison à un contrôle de contenu nommé](#add-and-bind-to-a-named-content-control)).

Lorsque vous utilisez des modèles avec votre application, que le complément réside dans le modèle au moment où l’utilisateur a créé le document ou que le complément insère un modèle, vous pouvez également incorporer d’autres éléments de l’API pour vous aider à créer une expérience plus interactive et robuste. Par exemple, vous pouvez inclure des données d’identification dans un composant customXML que vous pouvez utiliser pour déterminer le type de modèle, afin de fournir à l’utilisateur des options propres au modèle. Pour en savoir plus sur la façon de travailler avec customXML dans vos compléments, voir les ressources supplémentaires qui suivent.

## <a name="see-also"></a>Voir aussi

- [API JavaScript pour Office](../reference/javascript-api-for-office.md)
- [Norme ECMA-376 : Formats de fichier Office Open XML](https://www.ecma-international.org/publications/standards/Ecma-376.htm) (accéder ici au guide de langage complet et à la documentation correspondante sur Open XML)
- [Exploration de l’API JavaScript Office : liaison de données et parties XML personnalisées](/archive/msdn-magazine/2013/april/microsoft-office-exploring-the-javascript-api-for-office-data-binding-and-custom-xml-parts)
