---
ms.date: 06/20/2019
description: Créer des fonctions personnalisées dans Excel à l’aide de JavaScript.
title: Créer des fonctions personnalisées dans Excel
localization_priority: Priority
ms.openlocfilehash: e8f53919ebd5e44fe04e45dfd05192c77324f3aa
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127890"
---
# <a name="create-custom-functions-in-excel"></a>Créer des fonctions personnalisées dans Excel 

Les fonctions personnalisées permettent aux développeurs d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément. Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`. Cet article explique comment créer des fonctions personnalisées dans Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

L’image animée suivante montre votre classeur appelant une fonction que vous avez créée avec JavaScript ou Typescript. Dans cet exemple, la fonction personnalisée `=MYFUNCTION.SPHEREVOLUME` calcule le volume d’une sphère.

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

Le code suivant définit la fonction personnalisée `=MYFUNCTION.SPHEREVOLUME`.

```js
/**
 * Returns the volume of a sphere. 
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return Math.pow(radius, 3) * 4 * Math.PI / 3;
}
CustomFunctions.associate("SPHEREVOLUME", sphereVolume)
```

> [!NOTE]
> La section [problèmes connus](#known-issues)plus loin dans cet article indique les limitations en cours de fonctions personnalisées.

## <a name="how-a-custom-function-is-defined-in-code"></a>Comment une fonction personnalisée est définie dans le code

Si vous utilisez le [générateur de Yo Office](https://github.com/OfficeDev/generator-office) pour créer un projet de complément de fonctions personnalisées Excel, vous constaterez qu’il crée des fichiers qui contrôlent totalement vos fonctions, votre volet des tâches et votre complément. Nous allons vous concentrer sur les fichiers importants pour les fonctions personnalisées :

| File | Format de fichier | Description |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>ou<br/>**./src/functions/functions.ts** | JavaScript<br/>ou<br/>TypeScript | Contient le code qui définit les fonctions personnalisées. |
| **./src/functions/functions.html** | HTML | Fournit une référence&lt;script&gt; au fichier JavaScript qui définit les fonctions personnalisées. |
| **./manifest.xml** | XML | Spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers JavaScript et HTML qui figurent plus haut dans ce tableau. Répertorie également les emplacements des autres fichiers que votre complément pourrait utiliser, tels que les fichiers du volet des tâches et les fichiers de commande. |

### <a name="script-file"></a>Fichier de script

Le fichier de script (**./src/functions/functions.js** ou **./src/functions/functions.ts**) contient le code qui définit des fonctions personnalisées, des commentaires qui définissent la fonction, et associe les noms des fonctions personnalisées à des objets dans le fichier de métadonnées JSON.

Le code suivant définit la fonction personnalisée `add`. Les commentaires du code sont utilisés pour générer un fichier de métadonnées JSON décrivant la fonction personnalisée pour Excel. Le commentaire obligatoire `@customfunction` est déclaré en premier, pour indiquer qu’il s’agit d’une fonction personnalisée. Vous pouvez également constater que deux paramètres sont déclarés, `first` et `second`, qui sont suivis de leurs propriétés `description`. Enfin, une description `returns` est fournie. Pour plus d’informations sur les commentaires requis pour votre fonction personnalisée, voir [Créer des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md).

Le code suivant appelle également `CustomFunctions.associate("ADD", add)` pour associer la fonction `add()` avec son ID dans le fichier de métadonnées JSON `ADD`. Pour plus d’informations sur l’association de fonctions, voir [Meilleures pratiques des fonctions personnalisées](custom-functions-best-practices.md#associating-function-names-with-json-metadata).

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number.
 * @param second Second number.
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
```

Notez que le fichier **functions.html** qui régit le chargement du runtime de fonctions personnalisées doit créer un lien vers le CDN actuel pour les fonctions personnalisées. Les projets préparés avec la version actuelle du générateur Yo Office font référence au CDN correct. Si vous mettez à niveau un projet de fonction personnalisée de mars 2019 ou antérieur, vous devez copier le code ci-dessous dans la page ** functions.html**.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/beta/hosted/custom-functions-runtime.js" type="text/javascript"></script>
```

### <a name="manifest-file"></a>Fichier manifeste

Le fichier manifeste XML pour un complément qui définit les fonctions personnalisées (**./manifest.xml** du projet créé par le Générateur de Yo Office) spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers HTML, JavaScript et JSON. 

Le marquage XML suivant présente un exemple des éléments`<ExtensionPoint>` et `<Resources>` que vous devez inclure dans le manifeste d’un complément pour activer les fonctions personnalisées. Si vous utilisez le générateur de Yo Office, vos fichiers de fonction personnalisée générés contiennent un fichier manifeste plus complexe que vous pouvez comparer sur [ce dépôt Github](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml).

> [!NOTE] 
> Les URL spécifiées dans le fichier manifeste pour les fonctions personnalisées de fichiers HTML, JavaScript et JSON doivent avoir le même sous-domaine et être accessibles publiquement.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> Les fonctions dans Excel sont précédées par l’espace de noms spécifié dans votre fichier manifeste XML. L’espace de noms d’une fonction vient avant le nom de fonction et les deux sont séparés par un point. Par exemple, pour appeler la fonction `ADD42` dans la cellule de feuille de calcul Excel, vous saisiriez `=CONTOSO.ADD42`, car `CONTOSO` est l’espace de noms et `ADD42` est le nom de la fonction spécifié dans le fichier JSON. L’espace de noms est destiné à être utilisé comme identificateur de votre entreprise ou du complément. Un espace de noms ne peut contenir que des points et des caractères alphanumériques.

## <a name="coauthoring"></a>Co-création

Excel sur le web et Windows avec un abonnement Office 365 vous permettent de co-créer des documents et cette fonctionnalité est disponible avec les fonctions personnalisées. Si votre classeur utilise une fonction personnalisée, votre collègue sera invité à charger le complément de la fonction personnalisée. Quand vous avez tous les deux chargé le complément, la fonction personnalisée peut partager les résultats via la co-création.

Pour plus d’informations sur la co-création, voir [À propos de la co-création dans Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).

## <a name="known-issues"></a>Problèmes connus

Consulter les problèmes connus sur notre[repo GitHub Fonctions Excel Personnalisées](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## <a name="next-steps"></a>Étapes suivantes

Vous voulez essayer les fonctions personnalisées ? Consultez la documentation sur le [démarrage rapide de fonction personnalisée](../quickstarts/excel-custom-functions-quickstart.md) ou le [didacticiel sur les fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md). 

Un autre moyen simple d’essayer des fonctions personnalisées consiste à utiliser [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), un complément qui vous permet d’expérimenter des fonctions personnalisées directement dans Excel. Vous pouvez essayer de créer votre propre fonction personnalisée ou utiliser les exemples fournis.

Êtes-vous prêt à en apprendre davantage sur les capacités des fonctions personnalisées ? Découvrez une vue d’ensemble de l’[architecture des fonctions personnalisées](custom-functions-architecture.md).

## <a name="see-also"></a>Voir aussi 
* [Configuration requise de fonctions personnalisées](custom-functions-requirements.md)
* [Instructions d’attribution de noms](custom-functions-naming.md)
* [Meilleures pratiques](custom-functions-best-practices.md)
* [Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur](make-custom-functions-compatible-with-xll-udf.md)
