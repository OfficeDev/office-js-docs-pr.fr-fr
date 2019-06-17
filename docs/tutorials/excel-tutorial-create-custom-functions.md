---
title: Didacticiel de fonctions personnalisées Excel
description: Dans ce didacticiel, vous allez créer un complément Excel qui contient une fonction personnalisée qui effectue des calculs, requiert des données web ou lance un flux de données web.
ms.date: 06/15/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: c2eedee19cf4a612c83b7a45f7c5c5dc3b3f6937
ms.sourcegitcommit: e112a9b29376b1f574ee13b01c818131b2c7889d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/15/2019
ms.locfileid: "34997385"
---
# <a name="tutorial-create-custom-functions-in-excel"></a>Didacticiel : créer des fonctions personnalisées dans Excel

Les fonctions personnalisées vous permettent d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément. Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`. Vous pouvez créer des fonctions personnalisées qui effectuent des tâches simples comme des calculs personnalisés ou des tâches plus complexes telles que la diffusion en continu des données en temps réel à partir du web dans une feuille de calcul.

Dans ce didacticiel, vous allez :
> [!div class="checklist"]
> * Créer un complément de fonction personnalisée à l’aide la [Générateur Yeoman de compléments Office](https://www.npmjs.com/package/generator-office). 
> * Utiliser une fonction personnalisée prédéfinie pour effectuer un calcul simple.
> * Créer une fonction personnalisée qui demande les données à partir du web.
> * Créer une fonction personnalisée qui diffuse les données en temps réel à partir du web.

## <a name="prerequisites"></a>Conditions requises

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* Excel sur Windows (version 1810 ou ultérieure) ou Excel Online

## <a name="create-a-custom-functions-project"></a>Créer un projet de fonctions personnalisées

 Pour commencer, vous devez créer le projet de code pour créer votre complément de fonction personnalisée. Le [Générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) configurera votre projet avec certaines fonctions personnalisées prédéfinies que vous pouvez tester. Si vous avez déjà exécuté le démarrage rapide des fonctions personnalisées et généré un projet, continuez à utiliser ce projet et passez à [cette étape](#create-a-custom-function-that-requests-data-from-the-web) .

1. Exécutez la commande suivante, puis répondez aux invitations comme suit.
    
    ```command&nbsp;line
    yo office
    ```
    
    * **Sélectionnez un type de projet :** `Excel Custom Functions Add-in project`
    * **Sélectionnez un type de script :** `JavaScript`
    * **Comment souhaitez-vous nommer votre complément ?** `stock-ticker`

    ![Le générateur de yeoman pour les compléments Office vous invite pour les fonctions personnalisées](../images/UpdatedYoOfficePrompt.png)
    
    Le générateur crée le projet et installe les composants Node.js de la prise en charge.

2. Accédez au dossier racine du projet.
    
    ```command&nbsp;line
    cd stock-ticker
    ```

3. Créez le projet.
    
    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez. Si vous êtes invité à installer un certificat après avoir exécuté `npm run build`, acceptez d’installer le certificat fourni par le générateur Yeoman.

4. Démarrez le serveur web local qui est exécuté dans Node.js. Vous pouvez essayer le complément de fonction personnalisée dans Excel sur Windows ou Excel online.

# <a name="excel-on-windowstabexcel-windows"></a>[Excel sur Windows](#tab/excel-windows)

Pour tester votre complément dans Excel sous Windows, exécutez la commande suivante. Lorsque vous exécutez cette commande, le serveur Web local démarre et Excel s’ouvre avec votre complément chargé.

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

Pour tester votre complément dans Excel Online, exécutez la commande suivante. Lorsque vous exécutez cette commande, le serveur web local démarre.

```command&nbsp;line
npm run start:web
```

Pour utiliser votre complément de fonctions personnalisées, ouvrez un nouveau classeur dans Excel online. Dans ce classeur, effectuez les étapes suivantes pour chargement votre complément.

1. Dans Excel Online, sélectionnez l’onglet **Insérer**, puis **Compléments**.

   ![Insérer un ruban dans Excel Online avec l’icône mes compléments mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)
   
2. Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.

3. Sélectionnez **Parcourir... ** et accédez au répertoire racine du projet créé par le Générateur de Yo Office.

4. Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a>Essayer une fonction personnalisée prédéfinie

Le projet de fonctions personnalisées que vous avez créé contient des fonctions personnalisées prédéfinies, définies dans le fichier **./SRC/Functions/functions.js** . Le fichier**manifest.xml**indique que toutes les fonctions personnalisées appartiennent à l’`CONTOSO`espace de noms. L’espace de noms CONTOSO permet d’accéder aux fonctions personnalisées dans Excel.

Essayez de reproduire la`ADD` fonction personnalisée en complétant les étapes suivantes dans Excel:

1. Dans Excel, accédez à n’importe quelle cellule et entrez `=CONTOSO`. Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans l’espace de noms `CONTOSO`.

2. Exécutez la`CONTOSO.ADD` fonction, avec les nombres `10` et `200` comme paramètres d’entrée, en spécifiant la valeur`=CONTOSO.ADD(10,200)`suivante dans la cellule et appuyez sur entrée.

Le `ADD` fonction personnalisée calcule la somme des deux nombres que vous avez spécifiés et renvoie le résultat**210** .

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Créer une fonction personnalisée qui demande les données à partir du web

Intégration de données à partir du Web est un excellent moyen pour étendre Excel via les fonctions personnalisées. Vous allez ensuite créer une fonction personnalisée nommée `stockPrice` qui obtient des actions à partir d’une API Web et renvoie le résultat à la cellule d’une feuille de calcul. 

> [!NOTE]
> Le code suivant demande une cotation boursière à l’aide de l’API commerce IEX. Avant de pouvoir exécuter le code, vous devez [créer un compte gratuit avec le Cloud Iex](https://iexcloud.io/) afin que vous puissiez obtenir le jeton d’API requis dans la demande d’API.  

1. Dans le projet **boursier** , recherchez le fichier **./SRC/Functions/functions.js** et ouvrez-le dans votre éditeur de code.

2. Dans **functions. js**, recherchez `increment` la fonction et ajoutez le code suivant après cette fonction.

    ```js
    /**
    * Fetches current stock price
    * @customfunction 
    * @param {string} ticker Stock symbol
    * @returns {number} The current stock price.
    */
    function stockPrice(ticker) {
        //Note: In the following line, replace <YOUR_TOKEN_HERE> with the API token that you've obtained through your IEX Cloud account.
        var url = "https://cloud.iexapis.com/stable/stock/" + ticker + "/quote/latestPrice?token=<YOUR_TOKEN_HERE>"
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
    CustomFunctions.associate("STOCKPRICE", stockPrice);
    ```

    Le `CustomFunctions.associate` code associe le `id`de la fonction avec l’adresse de la fonction de `stockPrice` dans JavaScript afin qu’Excel peut appeler votre fonction.

3. Exécutez la commande suivante pour regénérer le projet.

    ```command&nbsp;line
    npm run build
    ```

4. Procédez comme suit (pour Excel sur Windows ou Excel Online) pour réenregistrer le complément dans Excel. Vous devez effectuer ces étapes avant que la nouvelle fonction ne soit disponible. 

# <a name="excel-on-windowstabexcel-windows"></a>[Excel sur Windows](#tab/excel-windows)

1. Fermez Excel, puis ouvrez de nouveau Excel.

2. Dans Excel, sélectionnez l’onglet **Insérer** , puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer un ruban dans Excel sur Windows avec la flèche mes compléments mise en surbrillance](../images/select-insert.png)

3. Dans la liste des compléments disponibles, recherchez la section **Compléments Développeur** et sélectionnez votre complément**bourse** pour effectuer cette opération.
    ![Insérer un ruban dans Excel sur Windows avec le complément de fonctions personnalisées Excel mis en surbrillance dans la liste mes compléments](../images/list-stock-ticker-red.png)

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

1. Dans Excel Online, sélectionnez l’onglet **insérer**, puis **compléments**. ![Insérer du ruban dans Excel Online avec l’icône Mes compléments mis en évidence](../images/excel-cf-online-register-add-in-1.png)

2. Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**. 

3. Sélectionnez **Parcourir... ** et accédez au répertoire racine du projet créé par le Générateur de Yo Office. 

4. Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.

---

<ol start="5">
<li> Essayez la nouvelle fonction. Dans la cellule <strong>B1</strong>, tapez le texte <strong>= CONTOSO. STOCKPRICE("MSFT")</strong> et appuyez sur ENTRÉE. Vous devriez voir que le résultat dans la cellule <strong>B1</strong> est le prix boursier actuel pour un partage de stock Microsoft.</li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a>Créer une fonction personnalisée asynchrone diffusion en continu

La fonction`stockPrice`que vous venez de créer renvoie le prix d’une action à un moment donné, mais les prix des actions changent constamment. Vous allez ensuite créer une fonction personnalisée nommée `stockPriceStream` qui obtient le prix d’une action chaque 1000 millisecondes.

1. Dans le projet **boursier** , ajoutez le code suivant à **./SRC/Functions/functions.js** et enregistrez le fichier.

    ```js
    /**
    * Streams real time stock price
    * @customfunction 
    * @param {string} ticker Stock symbol
    * @param {CustomFunctions.StreamingInvocation<number>} invocation
    */
    function stockPriceStream(ticker, invocation) {
        var updateFrequency = 1000 /* milliseconds*/;
        var isPending = false;

        var timer = setInterval(function() {
            // If there is already a pending request, skip this iteration:
            if (isPending) {
                return;
            }

            //Note: In the following line, replace <YOUR_TOKEN_HERE> with the API token that you've obtained through your IEX Cloud account.
            var url = "https://cloud.iexapis.com/stable/stock/" + ticker + "/quote/latestPrice?token=<YOUR_TOKEN_HERE>"
            isPending = true;

            fetch(url)
                .then(function(response) {
                    return response.text();
                })
                .then(function(text) {
                    invocation.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    invocation.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        invocation.onCanceled = () => {
            clearInterval(timer);
        };
    }
    CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
    ```
    
    Le `CustomFunctions.associate` code associe le `id`de la fonction avec l’adresse de la fonction de `stockPriceStream` dans JavaScript afin qu’Excel peut appeler votre fonction.
    
2. Exécutez la commande suivante pour regénérer le projet.

    ```command&nbsp;line
    npm run build
    ```

3. Procédez comme suit (pour Excel sur Windows ou Excel Online) pour réenregistrer le complément dans Excel. Vous devez effectuer ces étapes avant que la nouvelle fonction ne soit disponible. 

# <a name="excel-on-windowstabexcel-windows"></a>[Excel sur Windows](#tab/excel-windows)

1. Fermez Excel, puis ouvrez de nouveau Excel.

2. Dans Excel, sélectionnez l’onglet **Insérer** , puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer un ruban dans Excel sur Windows avec la flèche mes compléments mise en surbrillance](../images/select-insert.png)

3. Dans la liste des compléments disponibles, recherchez la section **Compléments Développeur** et sélectionnez votre complément**bourse** pour effectuer cette opération.
    ![Insérer un ruban dans Excel sur Windows avec le complément de fonctions personnalisées Excel mis en surbrillance dans la liste mes compléments](../images/list-stock-ticker-red.png)

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

Félicitations ! Vous avez créé un nouveau projet de fonctions personnalisées, essayé une fonction prédéfinie, créé une fonction personnalisée qui demande les données à partir du web et créé une fonction personnalisée qui diffuse les données en temps réel à partir du web. Vous pouvez également essayer de déboguer cette fonction à l’aide [des instructions de débogage de la fonction personnalisée](../excel/custom-functions-debugging.md). Pour en savoir plus sur les fonctions personnalisées dans Excel, passez à l’article suivant :

> [!div class="nextstepaction"]
> [Créer des fonctions personnalisées dans Excel](../excel/custom-functions-overview.md)

### <a name="legal-information"></a>Informations légales

Données fournies gratuitement par [IEX](https://iextrading.com/developer/). Afficher les [conditions d’utilisation de IEX](https://iextrading.com/api-exhibit-a/). L’utilisation de Microsoft de l’API IEX dans ce didacticiel est uniquement à des fins d’enseignement.
