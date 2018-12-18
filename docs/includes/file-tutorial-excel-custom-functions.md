# <a name="tutorial-create-custom-functions-in-excel"></a>Didacticiel : créer des fonctions personnalisées dans Excel

## <a name="introduction"></a>Présentation

Les fonctions personnalisées vous permettent d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément. Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`. Vous pouvez créer des fonctions personnalisées qui effectuent des tâches simples telles que des calculs personnalisés ou des tâches plus complexes telles que la diffusion en continu des données en temps réel à partir du web dans une feuille de calcul.

Dans ce didacticiel, vous allez :
> [!div class="checklist"]
> * Créer un projet de fonctions personnalisées à l’aide du générateur Yo Office
> * Utiliser une fonction personnalisée prédéfinie pour effectuer un calcul simple
> * Créer une fonction personnalisée qui demande les données à partir du web
> * Créer une fonction personnalisée qui diffuse les données en temps réel à partir du web

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a>Conditions préalables

* [Node.js et npm](https://nodejs.org/en/)

* [Git Bash](https://git-scm.com/downloads) (ou un autre client Git)

* La dernière version de [Yeoman](https://yeoman.io/) et le [générateur Yo Office](https://www.npmjs.com/package/generator-office). À l’invite de commandes, exécutez la commande suivante pour installer ces outils :

    ```bash
    npm install -g yo generator-office
    ```

* Excel pour Windows (1810 ou version ultérieure) ou Excel Online

* Rejoignez le[programme Office Insider](https://products.office.com/office-insider)(** Insider**niveau, anciennement appelé « Insider Fast »)

## <a name="create-a-custom-functions-project"></a>Créer un projet de fonctions personnalisées

Ce didacticiel commence à l’aide du générateur Yo Office pour créer les fichiers dont vous avez besoin pour votre projet fonctions personnalisées.

1. Exécutez la commande suivante, puis répondez aux invites comme suit.

    ```bash
    yo office
    ```

    * Choisissez un type de projet : `Excel Custom Functions Add-in project (...)`
    * Choisissez un type de script : `JavaScript`
    * Comment souhaitez-vous nommer votre complément ? `stock-ticker`

    ![Yo Office bash vous invite pour fonctions personnalisées](../images/yo-office-cfs-stock-ticker-3.png)

    Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de la prise en charge. Les fichiers de projet proviennent des référentiels [fonctions personnalisées Excel](https://github.com/OfficeDev/Excel-Custom-Functions)GitHub.

2. Accédez au dossier du projet.

    ```bash
    cd stock-ticker
    ```

3. Démarrez le serveur web local.

    * Si vous utilisez Excel pour Windows pour tester vos fonctions personnalisées, exécutez la commande suivante pour démarrer le serveur web local, ouvrir Excel et charger le complément :

        ```bash
        npm run start-desktop
        ```

    * Si vous utilisez Excel pour Windows pour tester vos fonctions personnalisées, exécutez la commande suivante pour démarrer le serveur web local : 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a>Essayer une fonction personnalisée prédéfinie

Le projet de fonctions personnalisées que vous avez crées en utilisant le générateur Yo Office contient certaines fonctions personnalisées prédéfinies, définies dans le fichier**src/functions/functions.js**. Le **manifest.xml** fichier dans le répertoire racine du projet indique que toutes les fonctions personnalisées appartiennent à l’ `CONTOSO` espace de noms.

Avant de pouvoir utiliser les fonctions personnalisées prédéfinies, vous devez inscrire le complément fonctions personnalisées dans Excel. Pour cela, complétez les étapes pour la plateforme que vous utiliserez dorénavant dans ce didacticiel.

* Si vous utilisez Excel pour Windows pour tester vos fonctions personnalisées :

    1. Dans Excel, sélectionnez l’onglet**insérer**, puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer du ruban dans Excel pour Windows avec la flèche Mes compléments mise en évidence](../images/excel-cf-register-add-in-1b.png)

    2. Dans la liste des compléments disponibles, recherchez la section **Compléments développeur** et sélectionnez le complément **Fonctions personnalisées Excel** pour l’enregistrer.
        ![Insérer du ruban dans Excel pour Windows avec le complément Fonctions personnalisées Excel mis en évidence dans la liste Mes compléments](../images/excel-cf-register-add-in-2.png)

* Si vous utilisez Excel Online pour tester vos fonctions personnalisées : 

    1. Dans Excel Online, sélectionnez l’onglet **insérer**, puis **compléments**. ![Insérer du ruban dans Excel Online avec l’icône Mes compléments mis en évidence](../images/excel-cf-online-register-add-in-1.png)

    2. Sélectionnez **Gérer mes compléments** et sélectionnez **Charger mon complément**. 

    3. Sélectionnez **Parcourir... ** et accédez au répertoire racine du projet créé par le Générateur de Yo Office. 

    4. Sélectionnez le fichier **manifest.xml** puis choisissez**Ouvrir**, puis sélectionnez **Charger**.

À ce stade, les fonctions personnalisées prédéfinies dans votre projet sont chargées et disponibles dans Excel. Essayez de reproduire la`ADD` fonction personnalisée en complétant les étapes suivantes dans Excel :

1. Dans une cellule, tapez **= CONTOSO**. Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans l’`CONTOSO` espace de noms.

2. Exécutez la`CONTOSO.ADD` fonction, avec les nombres `10` et `200` comme paramètres d’entrée, en spécifiant la valeur`=CONTOSO.ADD(10,200)`suivante dans la cellule et appuyez sur entrée.

Le `ADD` fonction personnalisée calcule la somme des deux nombres que vous avez spécifiés comme paramètres d’entrée. La saisie de`=CONTOSO.ADD(10,200)` doit générer le résultat **210** dans la cellule une fois que vous appuyez sur ENTRÉE.

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Créer une fonction personnalisée qui demande les données à partir du web

Que se passe-t-il si vous avez besoin d’une fonction qui peut demander le prix d’une action à partir d’une API et afficher le résultat dans la cellule d’une feuille de calcul ? Les fonctions personnalisées sont conçues de sorte que vous pouvez facilement demander les données à partir du web de façon asynchrone.

Procédez comme suit pour créer une fonction personnalisée nommée `stockPrice` qui accepte une action (par exemple, **MSFT**) et renvoie le prix de cette action. Cette fonction personnalisée utilise l’API de cotation IEX, qui est gratuit et ne requiert pas d’authentification.

1. Dans le projet **Bourse** que le Générateur de Yo Office a créé, recherchez le fichier**src/customfunctions.js** et ouvrez-le dans votre éditeur de code.

2. Ajoutez le code suivant à **customfunctions.js** et enregistrez le fichier.

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

3. Avant qu’Excel puisse rendre cette nouvelle fonction disponible aux utilisateurs finaux, vous devez spécifier les métadonnées qui décrivent cette fonction. Dans le projet**Bourse** que le Générateur de Yo Office a créé, recherchez le fichier**src/customfunctions.js** et ouvrez-le dans votre éditeur de code. Ajouter l’objet suivant à la`functions` matrice au sein du fichier **config/customfunctions.json** et enregistrez le fichier.

    JSON décrit la `stockPrice` fonction.

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

4. Vous devez réenregistrer le complément dans Excel afin que la nouvelle fonction soit disponible pour les utilisateurs finaux. Complétez les étapes pour la plateforme que vous utiliserez dorénavant dans ce didacticiel.

    * Si vous utilisez Excel pour Windows :

        1. Fermez Excel, puis ouvrez de nouveau Excel.

        2. Dans Excel, sélectionnez l’onglet**insérer**, puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer du ruban dans Excel pour Windows avec la flèche Mes compléments mise en évidence](../images/excel-cf-register-add-in-1b.png)

        1. Dans la liste des compléments disponibles, recherchez la section **Compléments développeur** et sélectionnez le complément **Fonctions personnalisées Excel** pour l’enregistrer.
            ![Insérer du ruban dans Excel pour Windows avec le complément Fonctions personnalisées Excel mis en évidence dans la liste Mes compléments](../images/excel-cf-register-add-in-2.png)

    * Si vous utilisez Excel Online : 

        1. Dans Excel Online, sélectionnez l’onglet **insérer**, puis **compléments**. ![Insérer du ruban dans Excel Online avec l’icône Mes compléments mis en évidence](../images/excel-cf-online-register-add-in-1.png)

        2. Sélectionnez **Gérer mes compléments** et sélectionnez **Charger mon complément**. 

        3. Sélectionnez **Parcourir... ** et accédez au répertoire racine du projet créé par le Générateur de Yo Office. 

        4. Sélectionnez le fichier **manifest.xml** puis choisissez**Ouvrir**, puis sélectionnez **Charger**.

5. À présent, nous allons essayer la nouvelle fonction. Dans la cellule **B1**, tapez le texte `=CONTOSO.STOCKPRICE("MSFT")` et appuyez sur ENTRÉE. Vous devriez voir que le résultat dans la cellule **B1** est le prix boursier actuel pour un partage de stock Microsoft.

## <a name="create-a-streaming-asynchronous-custom-function"></a>Créer une fonction personnalisée asynchrone diffusion en continu

La `stockPrice` fonction que vous venez de créer renvoie le prix d’une action à un moment donné, mais les prix des actions changent constamment. Nous allons créer une fonction personnalisée des flux de données à partir d’une API pour obtenir des mises à jour en temps réel sur un prix boursier.

Procédez comme suit pour créer une fonction personnalisée nommée `stockPriceStream` qui demande le prix d’une action boursière spécifique chaque 1000 millisecondes (à condition que la demande précédente soit terminée). Pendant la requête initiale en cours, vous pourrez afficher la valeur de l’espace réservé **## CHARGEMENT_DONNEES** la cellule dans laquelle la fonction est appelée. Lorsqu’une valeur est renvoyée par la fonction **## CHARGEMENT_DONNEES** sera remplacée par cette valeur dans la cellule.

1. Dans le projet**Bourse** que le Générateur de Yo Office a créé, ajoutez le fichier **src/customfunctions.js** et enregistrez le fichier.

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

2. Avant qu’Excel puisse rendre cette nouvelle fonction disponible aux utilisateurs finaux, vous devez spécifier les métadonnées qui décrivent cette fonction. Dans le projet**Bourse** que le Générateur de Yo Office a créé, ajoutez l’objet suivant à la `functions` matrice au sein du fichier**config/customfunctions.json** et enregistrez le fichier.

    JSON décrit la `stockPriceStream` fonction. Pour n’importe quelle fonction diffusion en continu, la `stream` propriété et la `cancelable` propriété doivent être définies `true` au sein de l’ `options` objet, comme illustré dans cet exemple de code.

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

3. Vous devez réenregistrer le complément dans Excel afin que la nouvelle fonction soit disponible pour les utilisateurs finaux. Complétez les étapes pour la plateforme que vous utiliserez dorénavant dans ce didacticiel.

    * Si vous utilisez Excel pour Windows :

        1. Fermez Excel, puis ouvrez de nouveau Excel.
        
        2. Dans Excel, sélectionnez l’onglet**insérer**, puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer du ruban dans Excel pour Windows avec la flèche Mes compléments mise en évidence](../images/excel-cf-register-add-in-1b.png)

        3. Dans la liste des compléments disponibles, recherchez la section **Compléments développeur** et sélectionnez le complément **Fonctions personnalisées Excel** pour l’enregistrer.
            ![Insérer du ruban dans Excel pour Windows avec le complément Fonctions personnalisées Excel mis en évidence dans la liste Mes compléments](../images/excel-cf-register-add-in-2.png)

    * Si vous utilisez Excel Online : 

        1. Dans Excel Online, sélectionnez l’onglet **insérer**, puis **compléments**. ![Insérer du ruban dans Excel Online avec l’icône Mes compléments mis en évidence](../images/excel-cf-online-register-add-in-1.png)

        2. Sélectionnez **Gérer mes compléments** et sélectionnez **Charger mon complément**. 

        3. Sélectionnez **Parcourir... ** et accédez au répertoire racine du projet créé par le Générateur de Yo Office. 

        4. Sélectionnez le fichier **manifest.xml** puis choisissez**Ouvrir**, puis sélectionnez **Charger**.

4. À présent, nous allons essayer la nouvelle fonction. Dans la cellule **C1**, tapez le texte `=CONTOSO.STOCKPRICESTREAM("MSFT")` et appuyez sur ENTRÉE. Si le marché est ouvert, vous devriez voir que le résultat dans la cellule **C1** constamment mis à jour pour refléter le prix en temps réel pour un partage d’actions Microsoft.

## <a name="next-steps"></a>Étapes suivantes

Dans ce didacticiel, vous avez créé un nouveau projet de fonctions personnalisées, essayé une fonction prédéfinie, créé une fonction personnalisée qui demande les données à partir du web et créé une fonction personnalisée qui diffuse les données en temps réel à partir du web. Pour en savoir plus sur les fonctions personnalisées dans Excel, passez à l’article suivant : 

> [!div class="nextstepaction"]
> [Créer des fonctions personnalisées dans Excel](../excel/custom-functions-overview.md)

## <a name="legal-information"></a>Informations légales

Données fournies gratuitement par [IEX](https://iextrading.com/developer/). Afficher les [conditions d’utilisation de IEX](https://iextrading.com/api-exhibit-a/). L’utilisation de Microsoft de l’API IEX dans ce didacticiel est uniquement à des fins d’enseignement.
