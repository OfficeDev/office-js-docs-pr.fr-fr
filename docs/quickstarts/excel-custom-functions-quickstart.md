---
ms.date: 01/13/2022
description: Développement de fonctions personnalisées dans le Guide de démarrage rapide d’Excel.
title: Démarrage rapide des fonctions personnalisées
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 2f15be99499e42c0598d70ae91b937c6637c0149
ms.sourcegitcommit: 45f7482d5adcb779a9672669360ca4d8d5c85207
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/19/2022
ms.locfileid: "62222303"
---
# <a name="get-started-developing-excel-custom-functions"></a>Prise en main du développement des fonctions personnalisées Excel

Avec les fonctions personnalisées, les développeurs peuvent désormais ajouter de nouvelles fonctions dans Excel en les définissant dans JavaScript ou Typescript dans le cadre d’un complément. Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`.

## <a name="prerequisites"></a>Conditions préalables

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Excel sur Windows (version 1904 ou ultérieure) ou Excel sur le Web.
- Les fonctions personnalisées d’Excel sont prises en charge dans Office sur Mac (connecté à un abonnement Office 365). Une mise à jour de ce didacticiel est bientôt prévue.

## <a name="build-your-first-custom-functions-project"></a>Créer votre premier projet de fonctions personnalisées

Pour commencer, vous utiliserez le Yeoman Générateur pour créer le projet de fonctions personnalisées. Cette option définit votre projet, avec la structure de dossiers correct, les fichiers source et les dépendances pour commencer le codage de vos fonctions personnalisées.

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Sélectionnez un type de projet :** `Excel Custom Functions Add-in project`
    - **Sélectionnez un type de script :** `JavaScript`
    - **Comment souhaitez-vous nommer votre complément ?** `starcount`

    ![Capture d’écran des invites d’interface de ligne de commande du générateur de compléments Yeoman Office pour les projets de fonctions personnalisées.](../images/starcountPrompt.png)

    Le générateur crée le projet et installe les composants Node.js de la prise en charge.

1. Le générateur Yeoman vous fournit des instructions dans votre ligne de commande sur la procédure à suivre pour le projet, mais ignorez-les et continuez de suivre nos instructions. Accédez au dossier racine du projet.

    ```command&nbsp;line
    cd starcount
    ```

1. Créez le projet.

    ```command&nbsp;line
    npm run build
    ```

1. Démarrez le serveur web local qui est exécuté dans Node.js. Vous pouvez tester le complément de fonction personnalisée dans Excel sur le web ou Windows. Vous serez peut-être invité à ouvrir le volet Office du complément, même si ce n’est pas obligatoire. Vous pouvez continuer à exécuter vos fonctions personnalisées sans ouvrir le volet Office de votre complément.

# <a name="excel-on-windows"></a>[Excel sur Windows](#tab/excel-windows)

Pour tester votre complément dans Excel sur Windows, exécutez la commande suivante. Lorsque vous exécutez cette commande, le serveur web local et Excel s’ouvrent avec votre complément chargé.

```command&nbsp;line
npm run start:desktop
```

[!INCLUDE [alert use https](../includes/alert-use-https.md)]
    
# <a name="excel-on-the-web"></a>[Excel sur le web](#tab/excel-online)

Pour tester votre complément dans Excel sur le web, exécutez la commande suivante. Lorsque vous exécutez cette commande, le serveur web local démarre. Remplacez « {url} » par l’URL d’un document Excel sur votre OneDrive ou une bibliothèque SharePoint sur laquelle vous disposez d’autorisations.

[!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

Pour utiliser votre complément de fonctions personnalisées, ouvrez un nouveau classeur dans Excel sur un navigateur. Dans ce classeur, procédez comme suit pour charger une version test de votre complément.

1. Dans Excel, sélectionnez l’onglet **Insertion**, puis **Compléments**.

   ![Capture d’écran du ruban Insertion dans Excel sur le web, avec le bouton Mes compléments mise en évidence.](../images/excel-cf-online-register-add-in-1.png)

1. Sélectionnez **Gérer mes Compléments** et sélectionnez **Télécharger mon complément**.

1. Sélectionnez **Parcourir...** et accédez au répertoire racine du projet créé par le Générateur de Yo Office.

1. Sélectionnez le fichier **manifest.xml** puis sélectionnez **Ouvrir**, puis sélectionnez **Télécharger**.

---

## <a name="try-out-a-prebuilt-custom-function"></a>Essayer une fonction personnalisée prédéfinie

Le projet de fonctions personnalisées que vous avez crées en utilisant le générateur Yeoman contient certaines fonctions personnalisées prédéfinies, définies dans le fichier **./src/functions/functions.js**. Le fichier **./manifest.xml** dans le répertoire racine du projet indique que toutes les fonctions personnalisées appartiennent à l’ `CONTOSO` espace de noms.

Dans votre classeur Excel, essayez la fonction personnalisée `ADD` en effectuant les étapes suivantes.

1. Sélectionnez une cellule et tapez `=CONTOSO`. Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans `CONTOSO` l'espace de noms.

1. Exécutez la`CONTOSO.ADD` fonction, en utilisant les nombres `10` et `200` comme paramètres d’entrée, en spécifiant la valeur`=CONTOSO.ADD(10,200)`suivante dans la cellule et appuyez sur entrée.

Le `ADD` fonction personnalisée calcule la somme des deux nombres que vous avez spécifiés comme paramètres d’entrée. La saisie de`=CONTOSO.ADD(10,200)` doit générer le résultat **210** dans la cellule une fois que vous appuyez sur ENTRÉE.

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé une fonction personnalisée dans un complément Excel ! Ensuite, créez un complément plus complexe avec la fonctionnalité de diffusion de données en continu. Le lien suivant vous guide tout au long des étapes suivantes dans le complément Excel avec le didacticiel de fonctions personnalisées.

> [!div class="nextstepaction"]
> [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web)

## <a name="troubleshooting"></a>Résolution des problèmes

Vous pouvez rencontrer des problèmes si vous exécutez le démarrage rapide plusieurs fois. Votre complément retourne une erreur lors de son chargement si le cache d'Office contient déjà une instance d'une fonction qui porte le même nom. Vous pouvez éviter cela en [vidant le cache Office ](../testing/clear-cache.md) avant d’exécuter `npm run start`.

:::image type="content" source="../images/custom-function-already-exists-error.png" alt-text="Message d’erreur Excel intitulé « Erreur lors de l’installation des fonctions ». Il contient le texte « Ce complément n’a pas été installé car une fonction personnalisée du même nom existe déjà ».":::

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble des fonctions personnalisées](../excel/custom-functions-overview.md)
- [Métadonnées fonctions personnalisées](../excel/custom-functions-json.md)
- [Exécution de fonctions personnalisées Excel](../excel/custom-functions-runtime.md)
