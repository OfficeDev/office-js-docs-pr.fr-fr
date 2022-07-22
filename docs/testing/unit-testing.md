---
title: Tests unitaires dans les compléments Office
description: Découvrez comment tester le code unitaire qui appelle les API JavaScript Office.
ms.date: 02/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 21858a68734ca5d07621f3e9c88b147ebac7dde6
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958747"
---
# <a name="unit-testing-in-office-add-ins"></a>Tests unitaires dans les compléments Office

Les tests unitaires vérifient les fonctionnalités de votre complément sans nécessiter de connexions réseau ou de service, y compris les connexions à l’application Office. Le code côté serveur de test unitaire, ainsi que le code côté client qui n’appelle *pas* les [API JavaScript Office](../develop/understanding-the-javascript-api-for-office.md), sont les mêmes dans les compléments Office que dans n’importe quelle application web. Il ne nécessite donc aucune documentation spéciale. Toutefois, le code côté client qui appelle les API JavaScript Office est difficile à tester. Pour résoudre ces problèmes, nous avons créé une bibliothèque pour simplifier la création d’objets Office fictifs dans les tests unitaires : [Office-Addin-Mock](https://www.npmjs.com/package/office-addin-mock). La bibliothèque facilite les tests des manières suivantes :

- Les API JavaScript Office doivent s’initialiser dans un contrôle webview dans le contexte d’une application Office (Excel, Word, etc.), afin qu’elles ne puissent pas être chargées dans le processus dans lequel les tests unitaires s’exécutent sur votre ordinateur de développement. La bibliothèque Office-Addin-Mock peut être importée dans vos fichiers de test, ce qui permet de simuler des API JavaScript Office dans le processus node.js dans lequel les tests s’exécutent.
- Les [API spécifiques à l’application](../develop/understanding-the-javascript-api-for-office.md#api-models) ont des méthodes de [chargement](../develop/application-specific-api-model.md#load) et [de synchronisation](../develop/application-specific-api-model.md#sync) qui doivent être appelées dans un ordre particulier par rapport à d’autres fonctions et les unes aux autres. En outre, la `load` méthode doit être appelée avec certains paramètres en fonction des propriétés des objets Office qui seront lues par le code *plus tard* dans la fonction testée. Toutefois, les frameworks de tests unitaires sont intrinsèquement sans état, de sorte qu’ils ne peuvent pas conserver d’enregistrement indiquant si `load` ou `sync` a été appelé ou quels paramètres ont été passés à `load`. Les objets fictifs que vous créez avec la bibliothèque Office-Addin-Mock ont un état interne qui effectue le suivi de ces éléments. Cela permet aux objets fictifs d’émuler le comportement d’erreur des objets Office réels. Par exemple, si la fonction testée tente de lire une propriété qui n’a pas été transmise pour `load`la première fois, le test retourne une erreur similaire à ce qu’Office retournerait.

La bibliothèque ne dépend pas des API JavaScript Office et peut être utilisée avec n’importe quelle infrastructure de test unitaire JavaScript, par exemple :

- [Jest](https://jestjs.io)
- [Moka](https://mochajs.org/)
- [Jasmin](https://jasmine.github.io/)

Les exemples de cet article utilisent l’infrastructure Jest. Il existe des exemples d’utilisation de l’infrastructure Mocha sur [la page d’accueil Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#examples).

## <a name="prerequisites"></a>Configuration requise

Cet article part du principe que vous êtes familiarisé avec les concepts de base du test unitaire et de la simulation, notamment la création et l’exécution de fichiers de test, et que vous avez une certaine expérience avec une infrastructure de test unitaire.

> [!TIP]
> Si vous travaillez avec Visual Studio, nous vous recommandons de lire l’article [Tests unitaires JavaScript et TypeScript dans Visual Studio](/visualstudio/javascript/unit-testing-javascript-with-visual-studio) pour obtenir des informations de base sur les tests unitaires JavaScript dans Visual Studio, puis revenir à cet article.

## <a name="install-the-tool"></a>Installer l’outil

Pour installer la bibliothèque, ouvrez une invite de commandes, accédez à la racine de votre projet de complément, puis entrez la commande suivante.

```command&nbsp;line
npm install office-addin-mock --save-dev
```

## <a name="basic-usage"></a>Utilisation de base

1. Votre projet aura un ou plusieurs fichiers de test. (Consultez les instructions de votre infrastructure de test et les exemples de fichiers de test dans Examples (#examples) ci-dessous.) Importez la bibliothèque, avec le mot clé ou `import` le `require` mot clé, dans n’importe quel fichier de test qui a un test d’une fonction qui appelle les API JavaScript Office, comme illustré dans l’exemple suivant.

   ```javascript
   const OfficeAddinMock = require("office-addin-mock");
   ```

1. Importez le module qui contient la fonction de complément que vous souhaitez tester avec le mot clé ou `import` le `require` mot clé. Voici un exemple qui suppose que votre fichier de test se trouve dans un sous-dossier du dossier avec les fichiers de code de votre complément.

   ```javascript
   const myOfficeAddinFeature = require("../my-office-add-in");
   ```

1. Créez un objet de données qui a les propriétés et sous-propriétés dont vous avez besoin pour tester la fonction. Voici un exemple d’objet qui simule la propriété Excel [Workbook.range.address](/javascript/api/excel/excel.range#excel-excel-range-address-member) et la méthode [Workbook.getSelectedRange](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedrange-member(1)) . Ce n’est pas le dernier objet fictif. Considérez-le comme un objet seed utilisé pour `OfficeMockObject` créer l’objet fictif final.

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

1. Transmettez l’objet de données au `OfficeMockObject` constructeur. Notez ce qui suit concernant l’objet retourné `OfficeMockObject` .

   - Il s’agit d’une simulation simplifiée d’un objet [OfficeExtension.ClientRequestContext](/javascript/api/office/officeextension.clientrequestcontext) .
   - L’objet fictif a tous les membres de l’objet de données et possède également des implémentations fictives des méthodes et `sync` des `load` méthodes.
   - L’objet fictif imite le comportement d’erreur crucial de l’objet `ClientRequestContext` . Par exemple, si l’API Office que vous testez tente de lire une propriété sans d’abord charger la propriété et appeler `sync`, le test échoue avec une erreur similaire à ce qui serait levé dans le runtime de production : « Erreur, propriété non chargée ».

   ```javascript
   const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);
   ```

    > [!NOTE]
    > La documentation de référence complète pour le `OfficeMockObject` type se trouve dans [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference).

1. Dans la syntaxe de votre infrastructure de test, ajoutez un test de la fonction. Utilisez l’objet `OfficeMockObject` à la place de l’objet qu’il simulacre, dans ce cas l’objet `ClientRequestContext` . L’exemple suivant se poursuit dans Jest. Cet exemple de test suppose que la fonction de complément testée est appelée `getSelectedRangeAddress`, qu’elle prend un `ClientRequestContext` objet comme paramètre et qu’elle est destinée à retourner l’adresse de la plage actuellement sélectionnée. L’exemple complet est [plus loin dans cet article](#mocking-a-clientrequestcontext-object).

   ```javascript
   test("getSelectedRangeAddress should return the address of the range", async function () {
     expect(await getSelectedRangeAddress(contextMock)).toBe("C2:G3");
   });
   ```

1. Exécutez le test conformément à la documentation de l’infrastructure de test et de vos outils de développement. En règle générale, il existe un fichier **package.json** avec un script qui exécute l’infrastructure de test. Par exemple, si Jest est l’infrastructure, **package.json** contient les éléments suivants :

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

## <a name="examples"></a>Exemples

Les exemples de cette section utilisent Jest avec ses paramètres par défaut. Ces paramètres prennent en charge les modules CommonJS. Consultez la [documentation Jest](https://jestjs.io/docs/getting-started) pour savoir comment configurer Jest et node.js pour prendre en charge les modules ECMAScript et prendre en charge TypeScript. Pour exécuter l’un de ces exemples, procédez comme suit.

1. Créez un projet de complément Office pour l’application hôte Office appropriée (par exemple, Excel ou Word). Une façon de procéder rapidement consiste à utiliser le [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md).
1. À la racine du projet, [installez Jest](https://jestjs.io/docs/getting-started).
1. [Installez l’outil office-addin-mock](#install-the-tool).
1. Créez un fichier exactement comme le premier fichier de l’exemple et ajoutez-le au dossier qui contient les autres fichiers sources du projet, souvent appelés `\src`.
1. Créez un sous-dossier dans le dossier du fichier source et donnez-lui un nom approprié, tel que `\tests`.
1. Créez un fichier exactement comme le fichier de test dans l’exemple et ajoutez-le au sous-dossier.
1. Ajoutez un `test` script au fichier **package.json** , puis exécutez le test, comme décrit dans [l’utilisation de base](#basic-usage).

### <a name="mocking-the-office-common-apis"></a>Simulation des API communes Office

Cet exemple suppose un complément Office pour n’importe quel hôte qui prend en charge les [API courantes Office](../develop/office-javascript-api-object-model.md) (par exemple, Excel, PowerPoint ou Word). Le complément a l’une de ses fonctionnalités dans un fichier nommé `my-common-api-add-in-feature.js`. Le code suivant montre le contenu du fichier. La `addHelloWorldText` fonction définit le texte « Hello World! » à tout ce qui est actuellement sélectionné dans le document ; par exemple ; une plage dans Word, une cellule dans Excel ou une zone de texte dans PowerPoint.

```javascript
const myCommonAPIAddinFeature = {

    addHelloWorldText: async () => {
        const options = { coercionType: Office.CoercionType.Text };
        await Office.context.document.setSelectedDataAsync("Hello World!", options);
    }
}
  
module.exports = myCommonAPIAddinFeature;
```

Le fichier de test, nommé `my-common-api-add-in-feature.test.js` est dans un sous-dossier, par rapport à l’emplacement du fichier de code du complément. Le code suivant montre le contenu du fichier. Notez que la propriété de niveau supérieur est `context`, un objet [Office.Context](/javascript/api/office/office.context) , de sorte que l’objet qui est simulé est le parent de cette propriété : un objet [Office](/javascript/api/office) . Tenez compte des informations suivantes à propos de ce code :

- Le `OfficeMockObject` constructeur n’ajoute *pas* toutes les classes d’énumération Office à l’objet fictif `Office` . Par conséquent, la `CoercionType.Text` valeur référencée dans la méthode de complément doit être ajoutée explicitement dans l’objet seed.
- Étant donné que la bibliothèque JavaScript Office n’est pas chargée dans le processus de nœud, l’objet `Office` référencé dans le code de complément doit être déclaré et initialisé.

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

### <a name="mocking-the-outlook-apis"></a>Simulation des API Outlook

Bien que strictement parlant, les API Outlook font partie du modèle d’API commune, elles ont une architecture spéciale qui est créée autour de l’objet [Mailbox](/javascript/api/outlook/office.mailbox) . Nous avons donc fourni un exemple distinct pour Outlook. Cet exemple suppose qu’Outlook possède l’une de ses fonctionnalités dans un fichier nommé `my-outlook-add-in-feature.js`. Le code suivant montre le contenu du fichier. La `addHelloWorldText` fonction définit le texte « Hello World! » vers ce qui est actuellement sélectionné dans la fenêtre de composition de message.

```javascript
const myOutlookAddinFeature = {

    addHelloWorldText: async () => {
        Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
      }
}

module.exports = myOutlookAddinFeature;
```

Le fichier de test, nommé `my-outlook-add-in-feature.test.js` est dans un sous-dossier, par rapport à l’emplacement du fichier de code du complément. Le code suivant montre le contenu du fichier. Notez que la propriété de niveau supérieur est `context`, un objet [Office.Context](/javascript/api/office/office.context) , de sorte que l’objet qui est simulé est le parent de cette propriété : un objet [Office](/javascript/api/office) . Tenez compte des informations suivantes à propos de ce code :

- La `host` propriété sur l’objet fictif est utilisée en interne par la bibliothèque fictive pour identifier l’application Office. Il est obligatoire pour Outlook. Il ne sert actuellement à rien d’autre application Office.
- Étant donné que la bibliothèque JavaScript Office n’est pas chargée dans le processus de nœud, l’objet `Office` référencé dans le code de complément doit être déclaré et initialisé.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myOutlookAddinFeature = require("../my-outlook-add-in-feature");

// Create the seed mock object.
const mockData = {
  // Identify the host to the mock library (required for Outlook).
  host: "outlook",
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

### <a name="mocking-the-office-application-specific-apis"></a>Simulation des API spécifiques à l’application Office

Lorsque vous testez des fonctions qui utilisent les API spécifiques à l’application, veillez à simuler le type d’objet approprié. Il existe deux options :

- Simuler un [OfficeExtension.ClientRequestObject](/javascript/api/office/officeextension.clientrequestcontext). Procédez comme suit lorsque la fonction testée remplit les deux conditions suivantes :

  - Il n’appelle pas *d’hôte*.`run` fonction, telle [qu’Excel.run](/javascript/api/excel#Excel_run_batch_).
  - Il ne fait référence à aucune autre propriété ou méthode directe d’un objet *Host* .

- Simuler un objet *Hôte* , tel [qu’Excel](/javascript/api/excel) ou [Word](/javascript/api/word). Effectuez cette opération lorsque l’option précédente n’est pas possible.

Les sous-sections ci-dessous présentent des exemples des deux types de tests.

#### <a name="mocking-a-clientrequestcontext-object"></a>Simulation d’un objet ClientRequestContext

Cet exemple suppose qu’un complément Excel possède l’une de ses fonctionnalités dans un fichier nommé `my-excel-add-in-feature.js`. Le code suivant montre le contenu du fichier. Notez qu’il `getSelectedRangeAddress` s’agit d’une méthode d’assistance appelée à l’intérieur du rappel passé à `Excel.run`.

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

Le fichier de test, nommé `my-excel-add-in-feature.test.js` est dans un sous-dossier, par rapport à l’emplacement du fichier de code du complément. Le code suivant montre le contenu du fichier. Notez que la propriété de niveau supérieur est `workbook`, de sorte que l’objet qui est simulé est le parent d’un `Excel.Workbook`: un `ClientRequestContext` objet.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myExcelAddinFeature = require("../my-excel-add-in-feature");

// Create the seed mock object.
const mockData = {
    workbook: {
      range: {
        address: "C2:G3",
      },
      // Mock the Workbook.getSelectedRange method.
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

Cet exemple suppose qu’un complément Word possède l’une de ses fonctionnalités dans un fichier nommé `my-word-add-in-feature.js`. Le code suivant montre le contenu du fichier.

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

Le fichier de test, nommé `my-word-add-in-feature.test.js` est dans un sous-dossier, par rapport à l’emplacement du fichier de code du complément. Le code suivant montre le contenu du fichier. Notez que la propriété de niveau supérieur est `context`, un `ClientRequestContext` objet, de sorte que l’objet qui est simulé est le parent de cette propriété : un `Word` objet. Tenez compte des informations suivantes à propos de ce code :

- Lorsque le `OfficeMockObject` constructeur crée l’objet fictif final, il s’assure que l’objet enfant `ClientRequestContext` a `sync` et `load` les méthodes.
- Le `OfficeMockObject` constructeur n’ajoute *pas* de `run` fonction à l’objet fictif `Word` . Il doit donc être ajouté explicitement dans l’objet seed.
- Le `OfficeMockObject` constructeur n’ajoute *pas* toutes les classes d’énumération Word à l’objet fictif `Word` . Par conséquent, la `InsertLocation.end` valeur référencée dans la méthode de complément doit être ajoutée explicitement dans l’objet seed.
- Étant donné que la bibliothèque JavaScript Office n’est pas chargée dans le processus de nœud, l’objet `Word` référencé dans le code de complément doit être déclaré et initialisé.

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
  // Mock the Word.run function.
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

> [!NOTE]
> La documentation de référence complète pour le `OfficeMockObject` type se trouve dans [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference).

## <a name="see-also"></a>Voir aussi

- [Point d’installation de la page npm Office-Addin-Mock](https://www.npmjs.com/package/office-addin-mock) . 
- Le dépôt open source est [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock).
- [Jest](https://jestjs.io)
- [Moka](https://mochajs.org/)
- [Jasmin](https://jasmine.github.io/)
