---
title: Test unitaire dans les Office de test
description: Découvrez comment unitér le code de test qui appelle Office API JavaScript
ms.date: 11/30/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8824b8e759e3c1acecf30683f2b89bb41bd558f3
ms.sourcegitcommit: 5daf91eb3be99c88b250348186189f4dc1270956
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/01/2021
ms.locfileid: "61242039"
---
# <a name="unit-testing-in-office-add-ins"></a>Test unitaire dans les Office de test

Les tests unitaires vérifient les fonctionnalités de votre Office sans nécessiter de connexions réseau ou de service. Le code côté serveur de test unitaire  et le code côté client qui n’appellent pas les API [JavaScript Office](../develop/understanding-the-javascript-api-for-office.md)sont les mêmes dans les applications Office que dans n’importe quelle application web, il ne nécessite donc aucune documentation spéciale. Toutefois, le code côté client qui appelle Office API JavaScript est difficile à tester. Pour résoudre ces problèmes, nous avons créé une bibliothèque pour simplifier la création d’objets Office facturants dans des tests unitaires : [Office-Addin-Mock](https://www.npmjs.com/package/office-addin-mock). La bibliothèque facilite les tests des manières suivantes :

- Les API JavaScript Office doivent s’initialiser dans un contrôle webview dans le contexte d’une application Office (Excel, Word, etc.), de sorte qu’elles ne peuvent pas être chargées dans le processus dans lequel les tests unitaires s’exécutent sur votre ordinateur de développement. La bibliothèque Office-Addin-Mock peut être importée dans vos fichiers de test, ce qui permet la simulation d’API JavaScript Office à l’intérieur du processus node.js dans lequel les tests s’exécutent.
- Les [API spécifiques à l’application](../develop/understanding-the-javascript-api-for-office.md#api-models) ont des méthodes de chargement et de synchronisation qui doivent être appelées dans un ordre particulier par rapport à d’autres fonctions et entre elles. [](../develop/application-specific-api-model.md#load) [](../develop/application-specific-api-model.md#sync) En outre, la méthode doit être appelée avec certains paramètres en fonction des propriétés des objets Office qui seront lues par code ultérieurement dans la fonction `load` testée.  Toutefois, les frameworks de test d’unité sont intrinsèquement sans état, de sorte qu’ils ne peuvent pas garder un enregistrement de l’appel ou de l’appel des paramètres qui ont `load` `sync` été passés à `load` . Les objets facturants que vous créez avec la bibliothèque Office-Addin-Mock ont un état interne qui assure le suivi de ces éléments. Cela permet aux objets facturables d’émuler le comportement d’erreur des objets Office réels. Par exemple, si la fonction en cours de test tente de lire une propriété qui n’a pas été passée pour la première fois, le test retourne une erreur semblable à ce que Office `load` renvoyait.

La bibliothèque ne dépend pas des API JavaScript Office et peut être utilisée avec n’importe quelle infrastructure de test d’unité JavaScript, telle que :

- [Jest](https://jestjs.io)
- [Mocha](https://mochajs.org/)
- [Tso](https://jasmine.github.io/)

Les exemples de cet article utilisent l’infrastructure Jest. Il existe des exemples d’utilisation de l’infrastructure Mocha Office page d’accueil [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#examples).

## <a name="prerequisites"></a>Conditions préalables

Cet article part du principe que vous connaissez les concepts de base des tests unitaires et de la maquette, notamment la création et l’utilisation de fichiers de test, et que vous avez une certaine expérience avec une infrastructure de test unitaire.

> [!TIP]
> Si vous travaillez avec Visual Studio, nous vous recommandons de lire l’article Unit [testing JavaScript et TypeScript](/visualstudio/javascript/unit-testing-javascript-with-visual-studio) dans Visual Studio pour obtenir des informations de base sur le test d’unité JavaScript dans Visual Studio, puis de revenir à cet article.

## <a name="install-the-tool"></a>Installer l’outil

Pour installer la bibliothèque, ouvrez une invite de commandes, accédez à la racine de votre projet de add-in, puis entrez la commande suivante.

```command&nbsp;line
npm install office-addin-mock --save-dev
```

## <a name="basic-usage"></a>Utilisation de base

1. Votre projet aura un ou plusieurs fichiers de test. (Consultez les instructions pour votre infrastructure de test et les exemples de fichiers de test dans Examples(#examples) ci-dessous.) Importez la bibliothèque, avec le ou le mot clé, dans n’importe quel fichier de test qui dispose d’un test d’une fonction qui appelle les API JavaScript Office, comme illustré dans l’exemple `require` `import` suivant.

   ```javascript
   const OfficeAddinMock = require("office-addin-mock");
   ```

1. Importez le module qui contient la fonction de module que vous souhaitez tester avec le mot clé `require` ou le `import` mot clé. L’exemple suivant suppose que votre fichier de test se trouve dans un sous-dossier du dossier contenant les fichiers de code de votre add-in.

   ```javascript
   const myOfficeAddinFeature = require("../my-office-add-in");
   ```

1. Créez un objet de données qui possède les propriétés et sous-propriétés dont vous avez besoin pour tester la fonction. Voici un exemple d’objet qui se Excel propriété [Workbook.range.address](/javascript/api/excel/excel.range#address) et la méthode [Workbook.getSelectedRange.](/javascript/api/excel/excel.workbook#getSelectedRange__) Il ne s’agit pas de l’objet mock final. Pensez-le comme un objet de début utilisé pour `OfficeMockObject` créer l’objet mock final.

   ```javascript
   const mockData = {
     workbook: {
       range: {
         address: "C2:G3",
       },
       getSelectedRange: function () {
         return this.range;
       },
     },
   };
   ```

1. Passez l’objet de données au `OfficeMockObject` constructeur. Notez ce qui suit sur l’objet `OfficeMockObject` renvoyé.

   - Il s’agit d’une maquette simplifiée d’un [objet OfficeExtension.ClientRequestContext.](/javascript/api/office/officeextension.clientrequestcontext)
   - L’objet mock a tous les membres de l’objet de données et possède également des implémentations de maquettes des méthodes `load` `sync` et des objets.
   - L’objet simulé imite le comportement d’erreur crucial de `ClientRequestContext` l’objet. Par exemple, si l’API Office que vous testez tente de lire une propriété sans avoir d’abord chargé la propriété et appelé, le test échouera avec une erreur semblable à celle qui serait lancée dans le runtime de production : « Erreur, propriété non chargée `sync` ».

   ```javascript
   const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);
   ```

    > [!NOTE]
    > La documentation de référence complète pour `OfficeMockObject` le type est [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference).

1. Dans la syntaxe de votre infrastructure de test, ajoutez un test de la fonction. Utilisez `OfficeMockObject` l’objet à la place de l’objet qu’il simula, dans ce cas `ClientRequestContext` l’objet. L’exemple suivant se poursuit dans Jest. Cet exemple de test suppose que la fonction de add-in en cours de test est appelée, qu’elle prend un objet comme paramètre et qu’elle est destinée à renvoyer l’adresse de la plage actuellement `getSelectedRangeAddress` `ClientRequestContext` sélectionnée. L’exemple complet est [plus loin dans cet article.](#mocking-a-clientrequestcontext-object)

   ```javascript
   test("getSelectedRangeAddress should return the address of the range", async function () {
     expect(await getSelectedRangeAddress(contextMock)).toBe("C2:G3");
   });
   ```

1. Exécutez le test conformément à la documentation de l’infrastructure de test et de vos outils de développement. En règle générale, il existe **un fichier package.json** avec un script qui exécute l’infrastructure de test. Par exemple, si Jest est l’infrastructure, **package.json** contient les éléments suivants :

   ```json
   "scripts": {
     "test": "jest",
     -- other scripts omitted --  
   }
   ```

   Pour exécuter le test, entrez ce qui suit dans une invite de commandes à la racine du projet.

   ```command&nbsp;line
   npm test
   ```

## <a name="examples"></a>範例

Les exemples de cette section utilisent Jest avec ses paramètres par défaut. Ces paramètres prise en charge les modules CommonJS. Consultez [la documentation de Jest](https://jestjs.io/docs/getting-started) pour savoir comment configurer Jest et node.js pour prendre en charge les modules ECMAScript et TypeScript. Pour exécuter l’un de ces exemples, exécutez les étapes suivantes.

1. Créez un Office pour l’application Office hôte appropriée (par exemple, Excel ou Word). Une façon de le faire rapidement consiste à utiliser [l’outil Yo Office](https://github.com/OfficeDev/generator-office).
1. À la racine du projet, [installez Jest](https://jestjs.io/docs/getting-started).
1. [Installez l’outil office-addin-mock.](#install-the-tool)
1. Créez un fichier exactement comme le premier fichier de l’exemple et ajoutez-le au dossier qui contient les autres fichiers sources du projet, souvent `\src` appelés .
1. Créez un sous-dossier dans le dossier de fichier source et donnez-lui un nom approprié, tel que `\tests` .
1. Créez un fichier exactement comme le fichier de test dans l’exemple et ajoutez-le au sous-dossier.
1. Ajoutez `test` un script au fichier **package.json,** puis exécutez le test, comme décrit dans [Utilisation de base.](#basic-usage)

### <a name="mocking-the-office-common-apis"></a>Mocking the Office Common APIs

Cet exemple suppose qu’un Office pour tout hôte qui prend en charge les API communes Office (par exemple, [Excel,](../develop/office-javascript-api-object-model.md) PowerPoint ou Word). Le add-in possède l’une de ses fonctionnalités dans un fichier nommé `my-common-api-add-in-feature.js` . L’exemple suivant montre le contenu du fichier. La `addHelloWorldText` fonction définit le texte « Hello World! » à tout ce qui est actuellement sélectionné dans le document ; par exemple ; une plage dans Word ou une cellule dans Excel, ou une zone de texte dans PowerPoint.

```javascript
const myCommonAPIAddinFeature = {

    addHelloWorldText: async () => {
        const options = { coercionType: Office.CoercionType.Text };
        await Office.context.document.setSelectedDataAsync("Hello World!", options);
    }
}
  
module.exports = myCommonAPIAddinFeature;
```

Le fichier de test, nommé, se trouve dans un sous-dossier, par rapport à l’emplacement du fichier de `my-common-api-add-in-feature.test.js` code du module. L’exemple suivant montre le contenu du fichier. Notez que la propriété de niveau supérieur `context` est un [Office. Objet](/javascript/api/office/office.context) de contexte, de sorte que l’objet qui est en cours de maquette est le parent de cette propriété : [Office](/javascript/api/office) objet. Tenez compte des informations suivantes à propos de ce code :

- Le constructeur n’ajoute pas toutes les classes d’enum Office à l’objet facturable. Par conséquent, la valeur référencé dans la méthode de add-in doit être ajoutée explicitement dans l’objet d’amorçage. `OfficeMockObject`  `Office` `CoercionType.Text`
- Étant donné que Office bibliothèque JavaScript n’est pas chargée dans le processus de nœud, l’objet référencé dans le code du module doit être déclaré et `Office` initialisé.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myCommonAPIAddinFeature = require("../my-common-api-add-in-feature");

// Create the seed mock object.
const mockData = {
    context: {
      document: {
        setSelectedDataAsync: function (data, options) {
          this.data = data;
          this.options = options;
        },
      },
    },
    // Mock the Office.CoercionType enum.
    CoercionType: {
      Text: {},
    },
};
  
// Create the final mock object from the seed object.
const officeMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Create the Office object that is called in the addHelloWorldText function.
global.Office = officeMock;

/* Code that calls the test framework goes below this line. */

// Jest test
test("Text of selection in document should be set to 'Hello World'", async function () {
    await myCommonAPIAddinFeature.addHelloWorldText();
    expect(officeMock.context.document.data).toBe("Hello World!");
});
```

### <a name="mocking-the-outlook-apis"></a>Mocking the Outlook APIs

Bien qu’à proprement parler, les API Outlook font partie du modèle API commun, elles ont une architecture spéciale qui est conçue autour de l’objet [Mailbox,](/javascript/api/outlook/office.mailbox) donc nous avons fourni un exemple distinct pour Outlook. Cet exemple suppose qu’un Outlook possède l’une de ses fonctionnalités dans un fichier nommé `my-outlook-add-in-feature.js` . L’exemple suivant montre le contenu du fichier. La `addHelloWorldText` fonction définit le texte « Hello World! » à tout ce qui est actuellement sélectionné dans la fenêtre de composition du message.

```javascript
const myOutlookAddinFeature = {

    addHelloWorldText: async () => {
        Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
      }
}

module.exports = myOutlookAddinFeature;
```

Le fichier de test, nommé, se trouve dans un sous-dossier, par rapport à l’emplacement du fichier de `my-outlook-add-in-feature.test.js` code du module. L’exemple suivant montre le contenu du fichier. Notez que la propriété de niveau supérieur `context` est un [Office. Objet](/javascript/api/office/office.context) de contexte, de sorte que l’objet qui est en cours de maquette est le parent de cette propriété : [Office](/javascript/api/office) objet. Tenez compte des informations suivantes à propos de ce code :

- Étant donné que Office bibliothèque JavaScript n’est pas chargée dans le processus de nœud, l’objet référencé dans le code du module doit être déclaré et `Office` initialisé.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myOutlookAddinFeature = require("../my-outlook-add-in-feature");

// Create the seed mock object.
const mockData = {
  context: {
    mailbox: {
      item: {
          setSelectedDataAsync: function (data) {
          this.data = data;
        },
      },
    },
  },
};
  
// Create the final mock object from the seed object.
const officeMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Create the Office object that is called in the addHelloWorldText function.
global.Office = officeMock;

/* Code that calls the test framework goes below this line. */

// Jest test
test("Text of selection in message should be set to 'Hello World'", async function () {
    await myOutlookAddinFeature.addHelloWorldText();
    expect(officeMock.context.mailbox.item.data).toBe("Hello World!");
});
```

### <a name="mocking-the-office-application-specific-apis"></a>Maquette des API Office’application

Lorsque vous testez des fonctions qui utilisent les API spécifiques à l’application, assurez-vous que vous faites une simulation du bon type d’objet. Il existe deux options :

- Mock a [OfficeExtension.ClientRequestObject](/javascript/api/office/officeextension.clientrequestcontext). Faites-le lorsque la fonction en cours de test répond aux deux conditions suivantes :

  - Il n’appelle pas un *hôte.*`run` par exemple, [Excel.run](/javascript/api/excel#Excel_run_batch_).
  - Il ne fait référence à aucune autre propriété ou méthode directe d’un *objet Hôte.*

- Mock a *Host* object, such as [Excel](/javascript/api/excel) or [Word](/javascript/api/word). Faites-le lorsque l’option précédente n’est pas possible.

Les sous-sections ci-dessous sont des exemples de ces deux types de tests.

#### <a name="mocking-a-clientrequestcontext-object"></a>Maquette d’un objet ClientRequestContext

Cet exemple suppose qu’un Excel qui possède l’une de ses fonctionnalités dans un fichier nommé `my-excel-add-in-feature.js` . L’exemple suivant montre le contenu du fichier. Notez `getSelectedRangeAddress` qu’il s’agit d’une méthode d’aide appelée à l’intérieur du rappel transmis à `Excel.run` .

```javascript
const myExcelAddinFeature = {
    
    getSelectedRangeAddress: async (context) => {
        const range = context.workbook.getSelectedRange();      
        range.load("address");

        await context.sync();
      
        return range.address;
    }
}

module.exports = myExcelAddinFeature;
```

Le fichier de test, nommé, se trouve dans un sous-dossier, par rapport à l’emplacement du fichier de `my-excel-add-in-feature.test.js` code du module. L’exemple suivant montre le contenu du fichier. Notez que la propriété de niveau supérieur est , donc l’objet qui est en cours de maquette est `workbook` le parent d’un `Excel.Workbook` : un `ClientRequestContext` objet.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myExcelAddinFeature = require("../my-excel-add-in-feature");

// Create the seed mock object.
const mockData = {
    workbook: {
      range: {
        address: "C2:G3",
      },
      // Mock the Workbook.getSelectRange method.
      getSelectedRange: function () {
        return this.range;
      },
    },
};

// Create the final mock object from the seed object.
const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);

/* Code that calls the test framework goes below this line. */

// Jest test
test("getSelectedRangeAddress should return address of selected range", async function () {
  expect(await myOfficeAddinFeature.getSelectedRangeAddress(contextMock)).toBe("C2:G3");
});
```

#### <a name="mocking-a-host-object"></a>Simulation d’un objet hôte

Cet exemple suppose qu’il s’agit d’un add-in Word qui possède l’une de ses fonctionnalités dans un fichier nommé `my-word-add-in-feature.js` . L’exemple suivant montre le contenu du fichier.

```javascript
const myWordAddinFeature = {

  insertBlueParagraph: async () => {
    return Word.run(async (context) => {
      // Insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
  
      // Change the font color to blue.
      paragraph.font.color = "blue";
  
      await context.sync();
    });
  }
}

module.exports = myWordAddinFeature;
```

Le fichier de test, nommé, se trouve dans un sous-dossier, par rapport à l’emplacement du fichier de `my-word-add-in-feature.test.js` code du module. L’exemple suivant montre le contenu du fichier. Notez que la propriété de niveau supérieur est , un objet, de sorte que l’objet qui est en cours de maquette est le parent de cette propriété `context` `ClientRequestContext` : un `Word` objet. Tenez compte des informations suivantes à propos de ce code :

- Lorsque le constructeur crée l’objet maquette final, il s’assure que l’objet `OfficeMockObject` enfant possède et les `ClientRequestContext` `sync` `load` méthodes.
- Le constructeur n’ajoute pas de méthode à l’objet facturable, il doit donc être ajouté explicitement `OfficeMockObject`  `run` dans `Word` l’objet d’amorçage.
- Le constructeur n’ajoute pas toutes les `OfficeMockObject` classes d’enum Word à l’objet facturable, donc la valeur référencé dans la méthode de  `Word` `InsertLocation.end` add-in doit être ajoutée explicitement dans l’objet d’amorçage.
- Étant donné que Office bibliothèque JavaScript n’est pas chargée dans le processus de nœud, l’objet référencé dans le code du module doit être déclaré et `Word` initialisé.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myWordAddinFeature = require("../my-word-add-in-feature");

// Create the seed mock object.
const mockData = {
  context: {
    document: {
      body: {
        paragraph: {
          font: {},
        },
        // Mock the Body.insertParagraph method.
        insertParagraph: function (paragraphText, insertLocation) {
          this.paragraph.text = paragraphText;
          this.paragraph.insertLocation = insertLocation;
          return this.paragraph;
        },
      },
    },
  },
  // Mock the Word.InsertLocation enum.
  InsertLocation: {
    end: "end",
  },
  // Mock the Word.run method.
  run: async function(callback) {
    await callback(this.context);
  },
};

// Create the final mock object from the seed object.
const wordMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Define and initialize the Word object that is called in the insertBlueParagraph function.
global.Word = wordMock;

/* Code that calls the test framework goes below this line. */

// Jest test set
describe("Insert blue paragraph at end tests", () => {

  test("color of paragraph", async function () {
    await myWordAddinFeature.insertBlueParagraph();  
    expect(wordMock.context.document.body.paragraph.font.color).toBe("blue");
  });

  test("text of paragraph", async function () {
    await myWordAddinFeature.insertBlueParagraph();
    expect(wordMock.context.document.body.paragraph.text).toBe("Hello World");
  });
})
```

## <a name="adding-mock-objects-properties-and-methods-dynamically-when-testing"></a>Ajout dynamique d’objets, de propriétés et de méthodes facturants lors des tests

Dans certains scénarios, des tests efficaces nécessitent la création ou la modification d’objets facturables au moment de l’utilisation . autrement dit, pendant que les tests sont en cours d’exécution. Les éléments suivants sont des exemples :

- La fonction testée se comporte différemment lorsqu’elle est appelée une deuxième fois. Vous devez d’abord tester la fonction avec un objet facturant, puis modifier cet objet et tester à nouveau la fonction avec l’objet maquette modifié.
- Vous devez tester une fonction sur plusieurs objets facturants similaires, mais pas identiques. Par exemple, vous devez tester une fonction avec un objet maquette qui possède une propriété de couleur, puis tester à nouveau la fonction avec un objet factuel qui possède une propriété de texte, mais qui est sinon identique à l’objet maquette d’origine.

Il `OfficeMockObject` dispose de trois méthodes pour vous aider dans ces scénarios.

- `OfficeMockObject.setMock` ajoute une propriété et une valeur à un `OfficeMockObject` objet. L’exemple suivant ajoute la `address` propriété.

    ```javascript
    rangeMock.setMock("address", "G6:K9");
    ```

- `OfficeMockObject.addMockFunction` ajoute une fonction de maquette à `OfficeMockObject` un objet, comme illustré dans l’exemple suivant.

    ```javascript
    workbookMock.addMockFunction("getSelectedRange", function () { 
      const range = {
        address: "B2:G5",
      };
      return range;
    });
    ```

    > [!NOTE]
    > Le paramètre de fonction est facultatif. Si elle n’est pas présente, une fonction vide est créée.

- `OfficeMockObject.addMock` ajoute un nouvel `OfficeMockObject` objet en tant que propriété à un objet existant et lui donne un nom. Il aurait le minimum de membres, `OfficeMockObject` tels que `load` et `sync` . Des membres supplémentaires peuvent être ajoutés avec les `setMock` méthodes `addMockFunction` et les méthodes. Voici un exemple qui ajoute un objet `Excel.WorkbookProtection` facturant en tant que propriété à un `protection` workbook factur. Il ajoute ensuite une `protected` propriété au nouvel objet mock.

    ```javascript
    workbookMock.addMock("protection");
    workbookMock.protection.setMock("protected", true);
    ```

> [!NOTE]
> La documentation de référence complète pour `OfficeMockObject` le type est [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference).

## <a name="see-also"></a>Voir aussi

- [Office d’installation de la page npm -Addin-Mock.](https://www.npmjs.com/package/office-addin-mock) 
- Le repo open source est [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock).
- [Jest](https://jestjs.io)
- [Mocha](https://mochajs.org/)
- [Tso](https://jasmine.github.io/)
