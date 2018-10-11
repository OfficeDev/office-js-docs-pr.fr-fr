# <a name="tutorial-create-custom-functions-in-excel"></a>Tutoriel : Créer des fonctions personnalisées dans Excel

## <a name="introduction"></a>Introduction

Les fonctions personnalisées vous permettent d’ajouter de nouvelles fonctions à Excel en définissant ces fonctions dans JavaScript comme partie d’un complément. Les utilisateurs dans Excel peuvent accéder aux fonctions personnalisées de la même façon qu’une fonction native dans Excel, telle que `SUM()`. Vous pouvez créer des fonctions personnalisées qui effectuent des tâches simples comme des calculs personnalisés ou des tâches plus complexes, comme la diffusion en continu des données en temps réel à partir du site Web dans une feuille de calcul.

Dans ce tutoriel, vous allez :
> [!div class="checklist"]
> * Créer un projet de fonctions personnalisées à l’aide du Générateur de Yo Office
> * Utiliser une fonction personnalisée prédéfinie pour effectuer un calcul simple
> * Créer une fonction personnalisée qui demande des données à partir du web
> * Créer une fonction personnalisée qui transmet les données en temps réel à partir du web

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a>Conditions préalables

* [Node.js et npm](https://nodejs.org/en/)

* [Git Bash](https://git-scm.com/downloads) (ou un autre client Git)

* La dernière version de [Yeoman](http://yeoman.io/) et le [Générateur de Yo Office](https://www.npmjs.com/package/generator-office). Pour installer ces outils globalement, exécutez la commande suivante par le biais de l’invite de commandes :

    ```bash
    npm install -g yo generator-office
    ```

* Excel pour Windows (numéro de build 10827 ou version ultérieure) ou Excel Online

* [Rejoindre le programme Office Insider](https://products.office.com/office-insider)

## <a name="create-a-custom-functions-project"></a>Créer un projet de fonctions personnalisées

Vous allez commencer ce tutoriel à l’aide du Générateur de Yo Office pour créer les fichiers dont vous avez besoin pour votre projet de fonctions personnalisées.

1. Exécutez la commande suivante, puis répondez aux invites comme suit.

    ```bash
    yo office
    ```

    * Choisissez un type de projet : `Excel Custom Functions Add-in project (...)`
    * Choisissez un type de script : `JavaScript`
    * Comment souhaitez-vous nommer votre complément ? `stock-ticker`

    ![Yo Office bash vous invite à fournir des fonctions personnalisées](../images/yo-office-cfs-stock-ticker-3.png)

    Après avoir exécuté l’assistant, le générateur crée les fichiers du projet et installe les composants Node de prise en charge. Les fichiers de projet viennent du référentiel GitHub [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) .

2. Accédez au dossier du projet.

    ```bash
    cd stock-ticker
    ```

3. Démarrez le serveur web local.

    * Si vous utilisez Excel pour Windows pour tester vos fonctions personnalisées, exécutez la commande suivante pour démarrer le serveur web local, lancer Excel et charger en parallèle le complément :

        ```bash
        npm start
        ```

    * Si vous allez utiliser Excel Online pour tester vos fonctions personnalisées, exécutez la commande suivante pour démarrer le serveur web local : 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a>Essayer une fonction personnalisée prédéfinie

Le projet de fonctions personnalisées que vous avez créé à l’aide du Générateur de Yo Office contient certaines fonctions personnalisées prédéfinies au niveau du fichier **src/customfunction.js**. Le fichier **manifest.xml** dans le répertoire racine du projet spécifie que toutes les fonctions personnalisées appartiennent à l'espace de noms `CONTOSO`.

Avant de pouvoir utiliser une des fonctions personnalisées prédéfinies, vous devez enregistrer le complément fonctions personnalisées dans Excel. Faites cela en procédant comme pour la plateforme que vous utiliserez dans ce tutoriel.

* Si vous utilisez Excel pour Windows pour tester vos fonctions personnalisées :

    1. Dans Excel, sélectionnez l’onglet **Insertion**, puis choisissez la flèche située à droite de **Mes applications**.  ![Insérez un ruban dans Excel pour Windows avec la flèche de Mes applications mise en surbrillance](../images/excel-cf-register-add-in-1b.png)

    2. Dans la liste des compléments disponibles, recherchez la section de **Compléments pour développeurs** et sélectionnez le complément **Fonctions personnalisées d'Excel** pour l’enregistrer.
        ![Insérez le ruban dans Excel pour Windows avec le complément des fonctions personnalisées d'Excel mis en surbrillance dans la liste du bouton Mes applications](../images/excel-cf-register-add-in-2.png)

* Si vous utilisez Excel Online pour tester vos fonctions personnalisées : 

    1. Dans Excel Online, choisissez l’onglet **Insertion** , puis choisissez **Compléments**.  ![Insérez le ruban dans Excel Online avec l'icône Mes applications mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)

    2. Sélectionnez **Gérer mes compléments** , sélectionnez **Télécharger mon complément**. 

    3. Cliquez sur **Parcourir** et accédez au répertoire racine du projet que le Générateur de Yo Office a créé. 

    4. Sélectionnez le fichier **manifest.xml** et choisissez **Ouvrir**, puis cliquez sur **Télécharger**.

À ce stade, les fonctions personnalisées prédéfinies dans votre projet sont chargés et disponibles dans Excel. Essayer la onction personnalisée `ADD` en effectuant les étapes suivantes dans Excel :

1. Dans une cellule, tapez **= CONTOSO**. Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions de le champ de noms pour `CONTOSO`.

2. Exécutez la fonction `CONTOSO.ADD`, avec les numéros `10` et `200` comme paramètres d’entrée, en spécifiant la valeur suivante dans la cellule et en appuyant sur, entrez :

    ```
    =CONTOSO.ADD(10,200)
    ```

La fonction personnalisée `ADD` calcule la somme de deux nombres que vous spécifiez comme paramètres d’entrée. Si vous tapez `=CONTOSO.ADD(10,200)`, vous devez obtenir le résultat **210** dans la cellule lorsque vous appuyez sur Entrée.

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Créer une fonction personnalisée qui demande des données à partir du web

Et si vous aviez besoin d’une fonction qui peut demander le prix d’une action à partir d’une API et afficher le résultat dans la cellule d’une feuille de calcul ? Les fonctions personnalisées sont conçues afin que vous puissiez aisément demander des données à partir du web de manière asynchrone.

Effectuez les étapes suivantes pour créer une fonction personnalisée nommée `stockPrice` qui a comme argument un symbole boursier (par exemple, **MSFT**) et renvoie le prix de l'action correspondante. Cette fonction personnalisée utilise l’API IEX Trading, qui est gratuite et ne nécessite pas d’authentification.

1. Dans le projet de **symboles boursiers** créé par le Générateur de Yo Office, recherchez le fichier **src/customfunctions.js** et ouvrez-le dans votre éditeur de code.

2. Ajoutez le code suivant à **customfunctions.js** et sauvegardez le fichier.

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

    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```

3. Avant qu'Excel ne puisse rendre cette nouvelle fonction disponible pour les utilisateurs finaux, vous devez spécifier les métadonnées décrivant cette fonction. Dans le projet de **symboles boursiers** créé par le Générateur de Yo Office, recherchez le fichier **config/customfunctions.js** et ouvrez-le dans votre éditeur de code. Ajoutez l’objet suivant au tableau `functions` dans le fichier **config/customfunctions.json** et sauvegardez le fichier.

    Ce code JSON décrit la fonction `stockPrice`.

    ```json
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
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

4. Vous devez réenregistrer le complément dans Excel pour que la nouvelle fonction soit disponible pour les utilisateurs finaux. Effectuez les étapes suivantes pour la plateforme que vous utilisez dans ce tutoriel.

    * Si vous utilisez Excel pour Windows :

        1. Fermez Excel, puis rouvrez Excel.

        2. Dans Excel, sélectionnez l’onglet **Insertion**, puis choisissez la flèche située à droite de **Mes applications**.  ![Insérez un ruban dans Excel pour Windows avec la flèche de Mes applications mise en surbrillance](../images/excel-cf-register-add-in-1b.png)

        1. Dans la liste des compléments disponibles, recherchez la section de **Compléments pour développeurs** et sélectionnez le complément **Fonctions personnalisées d'Excel** pour l’enregistrer.
            ![Insérez le ruban dans Excel pour Windows avec le complément des fonctions personnalisées d'Excel mis en surbrillance dans la liste du menu Mes applications](../images/excel-cf-register-add-in-2.png)

    * Si vous utilisez Excel Online : 

        1. Dans Excel Online, choisissez l’onglet **Insertion** , puis choisissez **Compléments**.  ![Insérez le ruban dans Excel Online avec l'icône Mes applications mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)

        2. Sélectionnez **Gérer mes compléments** , sélectionnez **Télécharger mon complément**. 

        3. Cliquez sur **Parcourir** et accédez au répertoire racine du projet que le Générateur de Yo Office a créé. 

        4. Sélectionnez le fichier **manifest.xml** et choisissez **Ouvrir**, puis cliquez sur **Télécharger**.

5. À présent, nous allons essayer la nouvelle fonction. Dans la cellule **B1**, tapez le texte `=CONTOSO.STOCKPRICE("MSFT")` et appuyez sur Entrée. Vous devriez voir que le résultat dans la cellule **B1** est le cours actuel d'une action Microsoft.

## <a name="create-a-streaming-asynchronous-custom-function"></a>Créer une fonction personnalisée asynchrone en continu

La fonction `stockPrice` que vous venez de créer renvoie le prix d’une action à un moment spécifique, mais les prix des actions varient constamment. Nous allons créer une fonction personnalisée qui récupère des flux de données à partir d’une API pour obtenir des mises à jour en temps réel sur les prix.

Effectuez les étapes suivantes pour créer une fonction personnalisée nommée `stockPriceStream` qui demande le prix de l'action toutes les 1000 millisecondes (à condition que la requête précédente soit terminée). Pendant que la requête initiale est en cours, vous pouvez voir le message d'indication **## GETTING_DATA** au niveau de la cellule dans laquelle la fonction est appelée. Lorsqu’une valeur est retournée par la fonction, **#GETTING_DATA** sera remplacé par cette valeur dans la cellule.

1. Dans le projet de **symboles boursiers** créé par le Générateur de Yo Office, ajoutez le code suivant à **src/customfunctions.js** et sauvegardez le fichier.

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

    CustomFunctionMappings.STOCKPRICESTREAM = stockPriceStream;
    ```

2. Avant qu'Excel ne puisse rendre cette nouvelle fonction disponible pour les utilisateurs finaux, vous devez spécifier les métadonnées décrivant cette fonction. Dans le projet de **symboles boursiers** créé par le Générateur de Yo Office, ajoutez l’objet suivant au tableau `functions` dans le fichier **config/customfunctions.json** et sauvegardez le fichier.

    Ce code JSON décrit la fonction `stockPriceStream`. Pour une fonction de diffusion en continu, la propriété `stream` et la propriété `cancelable` doivent être définies sur `true` dans l'objet `options`, comme illustré dans cet exemple de code.

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
                "description": "stock ticker name",
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

3. Vous devez réenregistrer le complément dans Excel pour que la nouvelle fonction soit disponible pour les utilisateurs finaux. Effectuez les étapes suivantes pour la plateforme que vous utilisez dans ce tutoriel.

    * Si vous utilisez Excel pour Windows :

        1. Fermez Excel, puis rouvrez Excel.
        
        2. Dans Excel, sélectionnez l’onglet **Insertion**, puis choisissez la flèche située à droite de **Mes applications**.  ![Insérez un ruban dans Excel pour Windows avec la flèche de Mes applications mise en surbrillance](../images/excel-cf-register-add-in-1b.png)

        3. Dans la liste des compléments disponibles, recherchez la section de **Compléments pour développeurs** et sélectionnez le complément **Fonctions personnalisées d'Excel** pour l’enregistrer.
            ![Insérez le ruban dans Excel pour Windows avec le complément des fonctions personnalisées d'Excel mis en surbrillance dans la liste du menu Mes applications](../images/excel-cf-register-add-in-2.png)

    * Si vous utilisez Excel Online : 

        1. Dans Excel Online, choisissez l’onglet **Insertion** , puis choisissez **Compléments**.  ![Insérez le ruban dans Excel Online avec l'icône Mes applications mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)

        2. Sélectionnez **Gérer mes compléments** , sélectionnez **Télécharger mon complément**. 

        3. Cliquez sur **Parcourir** et accédez au répertoire racine du projet que le Générateur de Yo Office a créé. 

        4. Sélectionnez le fichier **manifest.xml** et choisissez **Ouvrir**, puis cliquez sur **Télécharger**.

4. À présent, nous allons essayer la nouvelle fonction. Dans la cellule **C1**, tapez le texte `=CONTOSO.STOCKPRICESTREAM("MSFT")` et appuyez sur Entrée. À condition que le marché boursier est ouvert, vous devez voir que le résultat dans la cellule **C1** est constamment mis à jour pour refléter le prix en temps réel pour une action Microsoft.

## <a name="next-steps"></a>Étapes suivantes

Dans ce tutoriel, vous avez créé un nouveau projet de fonctions personnalisées, essayé une fonction prédéfinie, créé une fonction personnalisée qui demande des données à partir du web et créé une fonction personnalisée qui récupère des flux de données en temps réel à partir du web. Pour en savoir plus sur les fonctions personnalisées dans Excel, passez à l’article suivant : 

> [!div class="nextstepaction"]
> [Créer des fonctions personnalisées dans Excel](../excel/custom-functions-overview.md)

## <a name="legal-information"></a>Mentions légales

Données fournies gratuitement par [IEX](https://iextrading.com/developer/). Afficher les [conditions d’utilisation d’IEX](https://iextrading.com/api-exhibit-a/). L'utilisation de l’API IEX par Microsoft dans ce tutoriel est uniquement à des fins de formation.
