---
title: Didacticiel de fonctions personnalisées Excel
description: Dans ce didacticiel, vous allez créer un complément Excel qui contient une fonction personnalisée qui effectue des calculs, requiert des données web ou lance un flux de données web.
ms.date: 07/09/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: c7417c284beef787e35850ecbbb93b25ea5e1e87
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/13/2019
ms.locfileid: "36302608"
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

* Excel sur Windows (version 1904 ou ultérieure, connexion à l’abonnement Office 365) ou sur le Web

## <a name="create-a-custom-functions-project"></a>Créer un projet de fonctions personnalisées

 Pour commencer, vous devez créer le projet de code pour créer votre complément de fonction personnalisée. Le [Générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) configurera votre projet avec certaines fonctions personnalisées prédéfinies que vous pouvez tester. Si vous avez déjà exécuté le démarrage rapide des fonctions personnalisées et généré un projet, continuez à utiliser ce projet et passez à [cette étape](#create-a-custom-function-that-requests-data-from-the-web) .

[!include[note about Yeoman generator bug](../includes/note-yeoman-generator-bug-201908.md)]

1. Exécutez la commande suivante, puis répondez aux invitations comme suit.
    
    ```command&nbsp;line
    yo office
    ```
    
    * **Sélectionnez un type de projet :** `Excel Custom Functions Add-in project`
    * **Sélectionnez un type de script :** `JavaScript`
    * **Comment souhaitez-vous nommer votre complément ?** `starcount`

    ![Le générateur de yeoman pour les compléments Office vous invite pour les fonctions personnalisées](../images/starcountPrompt.png)
    
    Le générateur crée le projet et installe les composants Node.js de la prise en charge.

2. Accédez au dossier racine du projet.
    
    ```command&nbsp;line
    cd starcount
    ```

3. Créez le projet.
    
    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez. Si vous êtes invité à installer un certificat après avoir exécuté `npm run build`, acceptez d’installer le certificat fourni par le générateur Yeoman.

4. Démarrez le serveur web local qui est exécuté dans Node.js. Vous pouvez essayer le complément de fonction personnalisée dans Excel sur le Web ou Windows.

# <a name="excel-on-windows-or-mactabexcel-windows"></a>[Excel sur Windows ou Mac](#tab/excel-windows)

Pour tester votre complément dans Excel sous Windows ou Mac, exécutez la commande suivante: Lorsque vous exécutez cette commande, le serveur Web local démarre et Excel s’ouvre avec votre complément chargé.

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[Excel sur le Web](#tab/excel-online)

Pour tester votre complément dans Excel sur un navigateur, exécutez la commande suivante. Lorsque vous exécutez cette commande, le serveur web local démarre.

```command&nbsp;line
npm run start:web
```

Pour utiliser votre complément de fonctions personnalisées, ouvrez un nouveau classeur dans Excel sur le Web. Dans ce classeur, effectuez les étapes suivantes pour chargement votre complément.

1. Dans Excel, sélectionnez l’onglet **insertion** , puis **compléments**.

   ![Insérer un ruban dans Excel sur le Web avec l’icône mes compléments mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)
   
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

Intégration de données à partir du Web est un excellent moyen pour étendre Excel via les fonctions personnalisées. Ensuite, vous allez créer une fonction personnalisée `getStarCount` nommée qui indique le nombre d’étoiles dont dispose un référentiel GitHub donné.

1. Dans le projet **starcount** , recherchez le fichier **./SRC/Functions/functions.js** et ouvrez-le dans votre éditeur de code. 

2. Dans **Function. js**, ajoutez le code suivant: 

```JS
/**
  * Gets the star count for a given Github repository.
  * @customfunction 
  * @param {string} userName string name of Github user or organization.
  * @param {string} repoName string name of the Github repository.
  * @return {number} number of stars given to a Github repository.
  */
  async function getStarCount(userName, repoName) {
    try {
      //You can change this URL to any web request you want to work with.
      const url = "https://api.github.com/repos/" + userName + "/" + repoName;
      const response = await fetch(url);
      //Expect that status code is in 200-299 range
      if (!response.ok) {
        throw new Error(response.statusText)
      }
        const jsonResponse = await response.json();
        return jsonResponse.watchers_count;
    }
    catch (error) {
      return error;
    }
  }
```

3. Exécutez la commande suivante pour regénérer le projet.

    ```command&nbsp;line
    npm run build
    ```

4. Procédez comme suit (pour Excel sur le Web, Windows ou Mac) pour réenregistrer le complément dans Excel. Vous devez effectuer ces étapes avant que la nouvelle fonction ne soit disponible.

### <a name="excel-on-windows-or-mactabexcel-windows"></a>[Excel sur Windows ou Mac](#tab/excel-windows)

1. Fermez Excel, puis ouvrez de nouveau Excel.

2. Dans Excel, sélectionnez l’onglet **Insérer** , puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer un ruban dans Excel sur Windows avec la flèche mes compléments mise en surbrillance](../images/select-insert.png)

3. Dans la liste des compléments disponibles, recherchez la section **compléments pour développeurs** et sélectionnez le complément **starcount** pour l’enregistrer.
    ![Insérer un ruban dans Excel sur Windows avec le complément de fonctions personnalisées Excel mis en surbrillance dans la liste mes compléments](../images/list-starcount.png)


# <a name="excel-on-the-webtabexcel-online"></a>[Excel sur le Web](#tab/excel-online)

1. Dans Excel, sélectionnez l’onglet **insertion** , puis **compléments**.  ![Insérer un ruban dans Excel sur le Web avec l’icône mes compléments mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)

2. Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.

3. Sélectionnez **Parcourir... ** et accédez au répertoire racine du projet créé par le Générateur de Yo Office.

4. Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.

---

<ol start="5">
<li> Essayez la nouvelle fonction. Dans la cellule <strong>B1</strong>, tapez le texte <strong>= contoso. GETSTARCOUNT ("OfficeDev", "Excel-Custom-Functions")</strong> et appuyez sur entrée. Vous devriez voir que le résultat dans la cellule <strong>B1</strong> est le nombre actuel d’étoiles attribuées au [référentiel GitHub de fonctions personnalisées Excel](https://github.com/OfficeDev/Excel-Custom-Functions).</li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a>Créer une fonction personnalisée asynchrone diffusion en continu

La `getStarCount` fonction renvoie le nombre d’étoiles qu’un référentiel a à un moment donné. Les fonctions personnalisées peuvent également renvoyer des données qui changent en permanence. Ces fonctions sont appelées fonctions de diffusion en continu. Elles doivent inclure un `invocation` paramètre qui fait référence à la cellule à partir de laquelle la fonction a été appelée. Le `invocation` paramètre est utilisé pour mettre à jour le contenu de la cellule à tout moment.  

Dans l’exemple de code suivant, vous remarquerez qu’il existe deux `currentTime` fonctions `clock`, et. La `currentTime` fonction est une fonction statique qui n’utilise pas la diffusion en continu. Elle renvoie la date sous la forme d’une chaîne. La `clock` fonction utilise la `currentTime` fonction pour fournir la nouvelle fois toutes les secondes à une cellule dans Excel. Il utilise `invocation.setResult` pour fournir le temps à la cellule Excel et `invocation.onCanceled` pour gérer ce qui se produit lorsque la fonction est annulée.

1. Dans le projet **starcount** , ajoutez le code suivant à **./SRC/Functions/functions.js** et enregistrez le fichier.

```JS
/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}

 /**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
```

2. Exécutez la commande suivante pour regénérer le projet.

    ```command&nbsp;line
    npm run build
    ```

3. Procédez comme suit (pour Excel sur le Web, Windows ou Mac) pour réenregistrer le complément dans Excel. Vous devez effectuer ces étapes avant que la nouvelle fonction ne soit disponible. 

# <a name="excel-on-windows-or-mactabexcel-windows"></a>[Excel sur Windows ou Mac](#tab/excel-windows)

1. Fermez Excel, puis ouvrez de nouveau Excel.

2. Dans Excel, sélectionnez l’onglet **Insérer** , puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer un ruban dans Excel sur Windows avec la flèche mes compléments mise en surbrillance](../images/select-insert.png)

3. Dans la liste des compléments disponibles, recherchez la section **compléments pour développeurs** et sélectionnez le complément **starcount** pour l’enregistrer.
    ![Insérer un ruban dans Excel sur Windows avec le complément de fonctions personnalisées Excel mis en surbrillance dans la liste mes compléments](../images/list-starcount.png)

# <a name="excel-on-the-webtabexcel-online"></a>[Excel sur le Web](#tab/excel-online)

1. Dans Excel, sélectionnez l’onglet **insertion** , puis **compléments**.  ![Insérer un ruban dans Excel sur le Web avec l’icône mes compléments mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)

2. Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.

3. Sélectionnez **Parcourir... ** et accédez au répertoire racine du projet créé par le Générateur de Yo Office.

4. Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.

--- 

<ol start="4">
<li>Essayez la nouvelle fonction. Dans la cellule <strong>C1</strong>, tapez le texte <strong>= contoso. CLOCK ())</strong> , puis appuyez sur entrée. Vous devriez voir la date du jour, qui diffuse une mise à jour toutes les secondes. Bien que cette horloge constitue une seule horloge sur une boucle, vous pouvez utiliser la même idée de définir un minuteur sur des fonctions plus complexes qui effectuent des requêtes Web pour des données en temps réel.</li>
</ol>

## <a name="next-steps"></a>Étapes suivantes

Félicitations ! Vous avez créé un nouveau projet de fonctions personnalisées, testé une fonction prédéfinie, créé une fonction personnalisée qui demande des données à partir du Web et créé une fonction personnalisée qui diffuse les données. Vous pouvez également essayer de déboguer cette fonction à l’aide [des instructions de débogage de la fonction personnalisée](../excel/custom-functions-debugging.md). Pour en savoir plus sur les fonctions personnalisées dans Excel, passez à l’article suivant :

> [!div class="nextstepaction"]
> [Créer des fonctions personnalisées dans Excel](../excel/custom-functions-overview.md)
