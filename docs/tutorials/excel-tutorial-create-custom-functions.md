---
title: Didacticiel de fonctions personnalisées Excel (aperçu)
description: Dans ce didacticiel, vous allez créer un complément Excel qui contient une fonction personnalisée qui effectue des calculs, requiert des données web ou lance un flux de données web.
ms.date: 03/19/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 76f4d88b9da39a4d71927982836ee061b329a9b3
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451407"
---
# <a name="tutorial-create-custom-functions-in-excel-preview"></a>Didacticiel : créer des fonctions personnalisées dans Excel (aperçu)

Les fonctions personnalisées vous permettent d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément. Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`. Vous pouvez créer des fonctions personnalisées qui effectuent des tâches simples comme des calculs personnalisés ou des tâches plus complexes telles que la diffusion en continu des données en temps réel à partir du web dans une feuille de calcul.

Dans ce didacticiel, vous allez :
> [!div class="checklist"]
> * Créer un complément de fonction personnalisée à l’aide la [Générateur Yeoman de compléments Office](https://www.npmjs.com/package/generator-office). 
> * Utiliser une fonction personnalisée prédéfinie pour effectuer un calcul simple.
> * Créer une fonction personnalisée qui demande les données à partir du web.
> * Créer une fonction personnalisée qui diffuse les données en temps réel à partir du web.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a>Conditions requises

* [Node.js](https://nodejs.org/en/) (version 8.0.0 ou ultérieure)

* [Git Bash](https://git-scm.com/downloads) (ou un autre client Git)

* La dernière version de[Yeoman](https://yeoman.io/) et de [Yeoman Générateur de compléments Office](https://www.npmjs.com/package/generator-office). Pour installer ces outils globalement, exécutez la commande suivante à partir de l’invite de commande :

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > Même si vous avez précédemment installé la Yeoman générateur, nous vous recommandons une mise à jour de votre package à partir de la dernière version de npm.

* Excel pour Windows (version 64 bits 1810 ou ultérieure) ou Excel Online

* Rejoignez le[programme Office Insider](https://products.office.com/office-insider)(** niveau**Insider, anciennement appelé « Insider Fast »)

## <a name="create-a-custom-functions-project"></a>Créer un projet de fonctions personnalisées

 Pour commencer, vous devez créer le projet de code pour créer votre complément de fonction personnalisée. Le [ générateur Yeoman de compléments Office](https://www.npmjs.com/package/generator-office) permettront de configurer votre projet avec certaines fonctions personnalisées initiales que vous pouvez essayer.

1. Exécutez la commande suivante, puis répondez aux invitations comme suit.
    
    ```
    yo office
    ```
    
    * Choisissez un type de projet : `Excel Custom Functions Add-in project (...)`
    * Choisissez un type de script : `JavaScript`
    * Comment souhaitez-vous nommer votre complément ? `stock-ticker`
    
    ![Le générateur de yeoman pour les compléments Office vous invite pour les fonctions personnalisées](../images/12-10-fork-cf-pic.jpg)
    
    Le générateur Yeoman crée le projet et installe les composants Node.js de la prise en charge.

2. Accédez au dossier du projet.
    
    ```
    cd stock-ticker
    ```

3. Approuver le certificat auto-signé est nécessaire pour exécuter ce projet. Pour obtenir des instructions détaillées pour Windows ou Mac, voir [Ajout des Certificats Auto-signés comme Certificat Racine Approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).  

4. Construire le projet.
    
    ```
    npm run build
    ```

5. Démarrez le serveur web local qui est exécuté dans Node.js. Vous pouvez tester le complément de fonction personnalisée dans Excel pour Windows ou Excel Online.

# <a name="excel-for-windowstabexcel-windows"></a>[Excel pour Windows](#tab/excel-windows)

Exécutez la commande suivante.

```
npm start desktop
```

Cette commande démarre le serveur web et le complément sideloads de votre fonction personnalisée dans Excel pour Windows.

> [!NOTE]
> Si vous complément ne charge pas, vérifiez que vous avez correctement terminé l’étape 3. Vous pouvez également activer la journalisation de l' **[exécution](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** pour résoudre les problèmes liés au fichier manifeste XML de votre complément, ainsi que tous les problèmes d'installation ou d'exécution. La journalisation `console.log` de l'exécution écrit les instructions dans un fichier journal pour vous aider à trouver et à résoudre les problèmes.

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

Exécutez la commande suivante.

```
npm start web
```

Cette commande démarre le service web. Procédez comme suit pour votre complément sideload.

<ol type="a">
   <li>Dans Excel Online, sélectionnez l’onglet <strong>Insérer</strong>, puis <strong>Compléments</strong>.<br/>
   <img src="../images/excel-cf-online-register-add-in-1.png" alt="Insert ribbon in Excel Online with the My Add-ins icon highlighted"></li>
   <li>Sélectionnez<strong>Gérer mes Compléments</strong> et sélectionnez <strong>Télécharger mon complément</strong>.</li> 
   <li>Sélectionnez <strong>Parcourir... </strong> et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</li> 
   <li>Sélectionnez le fichier<strong>manifest.xml</strong> puis sélectionnez<strong>Ouvrir</strong>, puis sélectionnez <strong>Télécharger</strong>.</li>
</ol>

> [!NOTE]
> Si vous complément ne charge pas, vérifiez que vous avez correctement terminé l’étape 3.

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a>Essayer une fonction personnalisée prédéfinie

Le projet de fonctions personnalisées que vous avez créé déjà comporte deux fonctions personnalisées prédéfinies nommées AJOUTER et INCRÉMENT. Le code de ces fonctions prédéfinies se trouve dans le fichier **src/Functions/functions. js** . Le fichier**manifest.xml**indique que toutes les fonctions personnalisées appartiennent à l’`CONTOSO`espace de noms. L’espace de noms CONTOSO permet d’accéder aux fonctions personnalisées dans Excel.

Essayez de reproduire la`ADD` fonction personnalisée en complétant les étapes suivantes dans Excel:

1. Dans Excel, accédez à n’importe quelle cellule et entrez `=CONTOSO`. Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans l’espace de noms `CONTOSO`.

2. Exécutez la`CONTOSO.ADD` fonction, avec les nombres `10` et `200` comme paramètres d’entrée, en spécifiant la valeur`=CONTOSO.ADD(10,200)`suivante dans la cellule et appuyez sur entrée.

Le `ADD` fonction personnalisée calcule la somme des deux nombres que vous avez spécifiés et renvoie le résultat**210** .

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Créer une fonction personnalisée qui demande les données à partir du web

Intégration de données à partir du Web est un excellent moyen pour étendre Excel via les fonctions personnalisées. Vous allez ensuite créer une fonction personnalisée nommée `stockPrice` qui obtient des actions à partir d’une API Web et renvoie le résultat à la cellule d’une feuille de calcul. Cette fonction personnalisée utilise l’API de cotation IEX, qui est gratuit et ne requiert pas d’authentification.

1. Dans le projet **boursier** , recherchez le fichier **src/Functions/functions. js** et ouvrez-le dans votre éditeur de code.

2. Dans **functions. js**, recherchez `increment` la fonction et ajoutez le code suivant immédiatement après cette fonction.

    ```js
    function stockPrice(ticker) {
        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        return fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                return parseFloat(text);
            });

        // Note: in case of an error, the returned rejected Promise
        //    will be bubbled up to Excel to indicate an error.
    }

> [!NOTE]
> In the January Insiders 1901 Build, there is a bug preventing fetch calls from executing which will result in #VALUE!.
> To workaround this please use the [XMLHTTPRequest API](/office/dev/add-ins/excel/custom-functions-runtime#requesting-external-data) to make the web request.

3. In **functions.js**, locate the line `CustomFunctions.associate("INCREMENT", increment);`. Add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctions.associate("STOCKPRICE", stockprice);
    ```

    Le `CustomFunctions.associate` code associe le `id`de la fonction avec l’adresse de la fonction de `increment` dans JavaScript afin qu’Excel peut appeler votre fonction.

    Avant qu’Excel puisse utiliser votre fonction personnalisée, vous devez le décrire utilisant des métadonnées. Vous devez d’abord définir la méthode`id` utilisés dans le `associate`, ainsi que certaines autres métadonnées.


4. Ouvrez le fichier **src/Functions/functions. JSON** . Ajoutez l’objet JSON suivante à la matrice « fonctions » et enregistrez le fichier.

    ```JSON
    {
        "id": "STOCKPRICE",
        "name": "STOCKPRICE",
        "description": "Fetches current stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

    Cet objet JSON décrit la fonction`stockPrice`, ses paramètres, et le type de résultat qu’il renvoie.

5. Enregistrez de nouveau le complément dans Excel afin que la nouvelle fonction soit disponible. 

# <a name="excel-for-windowstabexcel-windows"></a>[Excel pour Windows](#tab/excel-windows)

1. Fermez Excel, puis ouvrez de nouveau Excel.

2. Dans Excel, sélectionnez l’onglet**insérer**, puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer du ruban dans Excel pour Windows avec la flèche Mes compléments mise en évidence](../images/excel-cf-register-add-in-1b.png)

3. Dans la liste des compléments disponibles, recherchez la section **Compléments Développeur** et sélectionnez votre complément**bourse** pour effectuer cette opération.
    ![Insérer du ruban dans Excel pour Windows avec le complément Fonctions personnalisées Excel mis en évidence dans la liste Mes compléments](../images/excel-cf-register-add-in-2.png)

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

1. Dans Excel Online, sélectionnez l’onglet **insérer**, puis **compléments**. ![Insérer du ruban dans Excel Online avec l’icône Mes compléments mis en évidence](../images/excel-cf-online-register-add-in-1.png)

2. Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**. 

3. Sélectionnez **Parcourir... ** et accédez au répertoire racine du projet créé par le Générateur de Yo Office. 

4. Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.

--- 

<ol start="6">
<li> Essayez la nouvelle fonction. Dans la cellule <strong>B1</strong>, tapez le texte <strong>= CONTOSO. STOCKPRICE("MSFT")</strong> et appuyez sur ENTRÉE. Vous devriez voir que le résultat dans la cellule <strong>B1</strong> est le prix boursier actuel pour un partage de stock Microsoft.</li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a>Créer une fonction personnalisée asynchrone diffusion en continu

La fonction`stockPrice`que vous venez de créer renvoie le prix d’une action à un moment donné, mais les prix des actions changent constamment. Vous allez ensuite créer une fonction personnalisée nommée `stockPriceStream` qui obtient le prix d’une action chaque 1000 millisecondes.

1. Dans le projet **bourse** , ajoutez le code suivant à **src/Functions/functions. js** et enregistrez le fichier.

    ```js
    function stockPriceStream(ticker, handler) {
        var updateFrequency = 1000 /* milliseconds*/;
        var isPending = false;

        var timer = setInterval(function() {
            // If there is already a pending request, skip this iteration:
            if (isPending) {
                return;
            }

            var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
            isPending = true;

            fetch(url)
                .then(function(response) {
                    return response.text();
                })
                .then(function(text) {
                    handler.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    handler.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        handler.onCanceled = () => {
            clearInterval(timer);
        };
    }
    
    CustomFunctions.associate("STOCKPRICESTREAM", stockpricestream);
    ```
    
    Avant qu’Excel puisse utiliser votre fonction personnalisée, vous devez le décrire utilisant des métadonnées.
    
2. Dans le projet **boursier** , ajoutez l'objet suivant à la `functions` matrice dans le fichier **src/Functions/functions. JSON** et enregistrez le fichier.
    
    ```json
    { 
        "id": "STOCKPRICESTREAM",
        "name": "STOCKPRICESTREAM",
        "description": "Streams real time stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ],
        "options": {
            "stream": true,
            "cancelable": true
        }
    }
    ```

    Cet élément JSON décrit la fonction`stockPriceStream`. Pour n’importe quelle fonction de diffusion en continu, la propriété`stream` et la propriété`cancelable`doivent être définies `true` au sein de l’ `options` objet, comme illustré dans cet exemple de code.

3. Enregistrez de nouveau le complément dans Excel afin que la nouvelle fonction soit disponible.

# <a name="excel-for-windowstabexcel-windows"></a>[Excel pour Windows](#tab/excel-windows)

1. Fermez Excel, puis ouvrez de nouveau Excel.

2. Dans Excel, sélectionnez l’onglet**insérer**, puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer du ruban dans Excel pour Windows avec la flèche Mes compléments mise en évidence](../images/excel-cf-register-add-in-1b.png)

3. Dans la liste des compléments disponibles, recherchez la section **Compléments Développeur** et sélectionnez votre complément**bourse** pour effectuer cette opération.
    ![Insérer du ruban dans Excel pour Windows avec le complément Fonctions personnalisées Excel mis en évidence dans la liste Mes compléments](../images/excel-cf-register-add-in-2.png)

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

1. Dans Excel Online, sélectionnez l’onglet **insérer**, puis **compléments**. ![Insérer du ruban dans Excel Online avec l’icône Mes compléments mis en évidence](../images/excel-cf-online-register-add-in-1.png)

2. Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.

3. Sélectionnez **Parcourir... ** et accédez au répertoire racine du projet créé par le Générateur de Yo Office.

4. Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.

--- 

<ol start="4">
<li>Essayez la nouvelle fonction. Dans la cellule <strong>C1</strong>, tapez le texte <strong>= CONTOSO. STOCKPRICE("MSFT")</strong> et appuyez sur ENTRÉE. Si le marché est ouvert, vous devriez voir que le résultat dans la cellule <strong>C1</strong> constamment mis à jour pour refléter le prix en temps réel pour un partage d’actions Microsoft.</li>
</ol>

## <a name="next-steps"></a>Étapes suivantes

Félicitations ! Vous avez créé un nouveau projet de fonctions personnalisées, essayé une fonction prédéfinie, créé une fonction personnalisée qui demande les données à partir du web et créé une fonction personnalisée qui diffuse les données en temps réel à partir du web. Vous pouvez également essayer de déboguer cette fonction à l'aide [des instructions de débogage de la fonction personnalisée](../excel/custom-functions-debugging.md). Pour en savoir plus sur les fonctions personnalisées dans Excel, passez à l’article suivant :

> [!div class="nextstepaction"]
> [Créer des fonctions personnalisées dans Excel](../excel/custom-functions-overview.md)

### <a name="legal-information"></a>Informations légales

Données fournies gratuitement par [IEX](https://iextrading.com/developer/). Afficher les [conditions d’utilisation de IEX](https://iextrading.com/api-exhibit-a/). L’utilisation de Microsoft de l’API IEX dans ce didacticiel est uniquement à des fins d’enseignement.


