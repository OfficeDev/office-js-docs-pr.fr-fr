---
title: Didacticiel de fonctions personnalisées Excel
description: Dans ce didacticiel, vous allez créer un complément Excel qui contient une fonction personnalisée qui effectue des calculs, requiert des données web ou lance un flux de données web.
ms.date: 10/08/2021
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 7f8a0cb7fcccce4861d77f23c0f3099fd1af2ec5
ms.sourcegitcommit: a37be80cf47a37c85b7f5cab216c160f4e905474
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2021
ms.locfileid: "60250454"
---
# <a name="tutorial-create-custom-functions-in-excel"></a>Didacticiel : créer des fonctions personnalisées dans Excel

Les fonctions personnalisées vous permettent d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément. Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`. Vous pouvez créer des fonctions personnalisées qui effectuent des tâches simples comme des calculs personnalisés ou des tâches plus complexes telles que la diffusion en continu des données en temps réel à partir du web dans une feuille de calcul.

Dans ce didacticiel, vous allez :
> [!div class="checklist"]
> - Créer un complément de fonction personnalisée à l’aide la [Générateur Yeoman de compléments Office](https://www.npmjs.com/package/generator-office). 
> - Utiliser une fonction personnalisée prédéfinie pour effectuer un calcul simple.
> - Créer une fonction personnalisée qui demande les données à partir du web.
> - Créer une fonction personnalisée qui diffuse les données en temps réel à partir du web.

## <a name="prerequisites"></a>Conditions préalables

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* Excel sur Windows (version 1904 ou ultérieure) ou Excel sur le Web.

## <a name="create-a-custom-functions-project"></a>Créer un projet de fonctions personnalisées

 Pour commencer, créez le projet de code pour créer votre complément de fonction personnalisé. Le [générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) configurera votre projet avec des fonctions personnalisées prédéfinies que vous pouvez essayer. Si vous avez déjà exécuté le démarrage rapide des fonctions personnalisées et généré un projet, continuez à utiliser ce projet et passez à [cette étape](#create-a-custom-function-that-requests-data-from-the-web) à la place.

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Sélectionnez un type de projet :** `Excel Custom Functions Add-in project`
    - **Sélectionnez un type de script :** `JavaScript`
    - **Comment souhaitez-vous nommer votre complément ?** `starcount`

    ![Capture d’écran des invites d’interface de ligne de commande du générateur de compléments Yeoman Office pour les projets de fonctions personnalisées.](../images/starcountPrompt.png)

    Le générateur crée le projet et installe les composants Node.js de la prise en charge.

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. Accédez au dossier racine du projet.

    ```command&nbsp;line
    cd starcount
    ```

1. Créez le projet.

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez. Si vous êtes invité à installer un certificat après avoir exécuté `npm run build`, acceptez d’installer le certificat fourni par le générateur Yeoman.

1. Démarrez le serveur web local qui est exécuté dans Node.js. Vous pouvez tester le complément de fonction personnalisée dans Excel sur le web ou sur Windows.

# <a name="excel-on-windows-or-mac"></a>[Excel sur Windows ou Mac](#tab/excel-windows)

Pour tester votre complément dans Excel sur Windows ou Mac, exécutez la commande suivante. Lorsque vous exécutez cette commande, le serveur web local et Excel s’ouvrent avec votre complément chargé.

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-web"></a>[Excel sur le web](#tab/excel-online)

Pour tester votre complément dans Excel sur un navigateur, exécutez la commande suivante. Lorsque vous exécutez cette commande, le serveur web local démarre.

```command&nbsp;line
npm run start:web
```

Pour utiliser votre complément de fonctions personnalisées, ouvrez un nouveau classeur dans Excel sur le web. Dans ce classeur, chargez une version test de votre complément en procédant comme suit.

1. Dans Excel, sélectionnez l’onglet **Insertion**, puis **Compléments**.

   ![Capture d’écran du ruban Insertion dans Excel sur le web, avec le bouton Mes compléments mise en évidence.](../images/excel-cf-online-register-add-in-1.png)

1. Sélectionnez **Gérer mes Compléments** et sélectionnez **Télécharger mon complément**.

1. Sélectionnez **Parcourir...** et accédez au répertoire racine du projet créé par le Générateur de Yo Office.

1. Sélectionnez le fichier **manifest.xml** puis sélectionnez **Ouvrir**, puis sélectionnez **Télécharger**.

---

## <a name="try-out-a-prebuilt-custom-function"></a>Essayer une fonction personnalisée prédéfinie

Le projet de fonctions personnalisées que vous avez créé contient certaines fonctions personnalisées prédéfinies, définies dans le fichier **./src/functions/functions.js**. Le fichier **manifest.xml** indique que toutes les fonctions personnalisées appartiennent à l’`CONTOSO`espace de noms. L’espace de noms CONTOSO permet d’accéder aux fonctions personnalisées dans Excel.

Ensuite, essayez la `ADD` fonction personnalisée en effectuant les étapes suivantes.

1. Dans Excel, accédez à n’importe quelle cellule et entrez `=CONTOSO`. Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans l’espace de noms `CONTOSO`.

1. Exécutez la`CONTOSO.ADD` fonction, avec les nombres `10` et `200` comme paramètres d’entrée, en spécifiant la valeur`=CONTOSO.ADD(10,200)`suivante dans la cellule et appuyez sur entrée.

Le `ADD` fonction personnalisée calcule la somme des deux nombres que vous avez spécifiés et renvoie le résultat **210** .

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Créer une fonction personnalisée qui demande les données à partir du web

Intégration de données à partir du Web est un excellent moyen pour étendre Excel via les fonctions personnalisées. Vous allez ensuite créer une fonction personnalisée nommée `getStarCount` qui affiche le nombre d’étoiles attribuées à un référentiel GitHub donné.

1. Dans le projet **starcount**, recherchez le fichier **./src/functions/functions.js** et ouvrez-le dans votre éditeur de code.

1. Dans le fichier **function.js**, ajoutez le code suivant.

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

1. Exécutez la commande suivante pour régénérer le projet.

    ```command&nbsp;line
    npm run build
    ```

1. Enregistrez de nouveau le complément dans Excel en procédant comme suit (pour Excel sur le web, Windows ou Mac). Vous devez suivre ces étapes pour que la nouvelle fonction devienne disponible.

### <a name="excel-on-windows-or-mac"></a>[Excel sur Windows ou Mac](#tab/excel-windows)

1. Fermez Excel, puis rouvrez-le.

1. Dans Excel, sélectionnez l’onglet **insérer**, puis cliquez sur la flèche vers le bas située à droite de **Mes compléments**.  ![Capture d’écran du ruban Insérer dans Excel sur Windows, avec la flèche vers le bas de Mes compléments mise en évidence.](../images/select-insert.png)

1. Dans la liste des compléments disponibles, recherchez la section **Compléments de développeur**, puis sélectionnez le complément **starcount** pour effectuer cette opération.
    ![Capture d’écran du ruban Insérer dans Excel sous Windows, avec le complément Fonctions personnalisées Excel mis en évidence dans la liste Mes compléments.](../images/list-starcount.png)

# <a name="excel-on-the-web"></a>[Excel sur le web](#tab/excel-online)

1. Dans Excel, sélectionnez l’onglet **Insertion**, puis sélectionnez **Compléments**.  ![Capture d’écran du ruban insérer dans Excel sur le Web, avec le bouton Mes compléments en surbrillance.](../images/excel-cf-online-register-add-in-1.png)

1. Sélectionnez **Gérer mes Compléments** et sélectionnez **Télécharger mon complément**.

1. Sélectionnez **Parcourir...** et accédez au répertoire racine du projet créé par le Générateur de Yo Office.

1. Sélectionnez le fichier **manifest.xml** puis sélectionnez **Ouvrir**, puis sélectionnez **Télécharger**.

5. Essayez la nouvelle fonction. Dans la cellule **B1**, tapez le texte **=CONTOSO. GETSTARCOUNT(« OfficeDev », « Excel-Custom-Functions »)**, puis appuyez sur Entrée. Le résultat dans la cellule **B1** doit correspondre au nombre d’étoiles actuellement attribuées au [référentiel GitHub Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions).

---

## <a name="create-a-streaming-asynchronous-custom-function"></a>Créer une fonction personnalisée asynchrone de diffusion en continu

La fonction `getStarCount` renvoie le nombre d’étoiles attribuées à un référentiel à un moment donné. Les fonctions personnalisées peuvent également renvoyer des données qui changent continuellement. Ces fonctions sont appelées fonctions de diffusion en continu. Elles doivent inclure un `invocation` paramètre qui fait référence à la cellule à partir de laquelle la fonction a été appelée. Le `invocation` paramètre permet de mettre à jour le contenu de la cellule à tout moment.  

Vous remarquerez que l’exemple de code suivant inclut deux fonctions (`currentTime` et `clock`). `currentTime` est une fonction statique qui n’utilise pas la diffusion en continu. Elle renvoie la date sous la forme d’une chaîne. La fonction `clock` utilise la fonction `currentTime` pour fournir la nouvelle heure toutes les secondes à une cellule dans Excel. Elle utilise `invocation.setResult` pour communiquer l’heure à la cellule Excel et `invocation.onCanceled` pour gérer le résultat de l’annulation de la fonction. 

Le projet **starcount** contient déjà les deux fonctions suivantes dans le fichier **./src/functions/functions.js** .

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

Pour tester les fonctions, tapez le texte **=CONTOSO. CLOCK()** dans la cellule **C1**, ensuite appuyez sur Entrée. La date du jour doit apparaître. Elle est mise à jour toutes les secondes. Cette horloge n’est qu’une minuterie incluse dans une boucle, mais vous pouvez vous inspirer de cette idée pour créer des fonctions plus complexes qui récupèrent des données en temps réel en exécutant des requêtes web.

## <a name="next-steps"></a>Étapes suivantes

Félicitations ! Vous avez créé un nouveau projet de fonctions personnalisées, essayé une fonction prédéfinie, créé une fonction personnalisée qui récupère des données à partir du web et créé une fonction personnalisée qui diffuse des données. Ensuite, vous pouvez modifier votre projet pour utiliser un runtime partagé, ce qui permet à votre fonction d’interagir plus facilement avec le volet de tâches. Suivez la procédure décrite dans cet article.

> [!div class="nextstepaction"]
> [Configurer votre complément pour utiliser un runtime partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
