---
ms.date: 06/20/2019
description: Développement de fonctions personnalisées dans le Guide de démarrage rapide d’Excel.
title: Démarrage rapide des fonctions personnalisées
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 8e7fbf247df04a12c38ad24d9ba6335a7f7bdaf8
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128579"
---
# <a name="get-started-developing-excel-custom-functions"></a>Prise en main du développement de fonctions personnalisées Excel

Avec les fonctions personnalisées, les développeurs peuvent désormais ajouter de nouvelles fonctions à Excel en les définissant en JavaScript ou en une machine à écrire dans le cadre d’un complément. Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe `SUM()`quelle fonction native dans Excel, comme.

## <a name="prerequisites"></a>Conditions préalables

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* Excel sur Windows (version 1904 ou ultérieure, connexion à l’abonnement Office 365) ou Excel sur le Web
* Les fonctions personnalisées d’Excel sont prises en charge dans Office sur Mac (connexion à l’abonnement Office 365) et une mise à jour de ce didacticiel est prochainement prévue.

>[!NOTE]
>Les fonctions personnalisées d’Excel ne sont pas prises en charge dans Office 2019 (achat unique).

## <a name="build-your-first-custom-functions-project"></a>Création de votre premier projet de fonctions personnalisées

Pour commencer, vous utiliserez le Yeoman Générateur pour créer le projet de fonctions personnalisées. Cette option définit votre projet, avec la structure de dossiers correct, les fichiers source et les dépendances pour commencer le codage de vos fonctions personnalisées.

1. Dans un dossier de votre choix, exécutez la commande suivante, puis répondez aux invites comme suit.

    ```command&nbsp;line
    yo office
    ```

    - **Sélectionnez un type de projet :** `Excel Custom Functions Add-in project`
    - **Sélectionnez un type de script :** `JavaScript`
    - **Comment souhaitez-vous nommer votre complément ?** `stock-ticker`

    ![Le générateur de yeoman pour les compléments Office vous invite pour les fonctions personnalisées](../images/UpdatedYoOfficePrompt.png)

    Le générateur crée le projet et installe les composants Node.js de la prise en charge.

2. Le générateur Yeoman vous donne des instructions dans votre ligne de commande sur ce qu’il faut faire du projet, mais il les ignore et continue de suivre nos instructions. Accédez au dossier racine du projet.

    ```command&nbsp;line
    cd stock-ticker
    ```

3. Créez le projet. 

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez. Si vous êtes invité à installer un certificat après avoir exécuté `npm run build`, acceptez d’installer le certificat fourni par le générateur Yeoman.

4. Démarrez le serveur web local qui est exécuté dans Node.js. Vous pouvez essayer le complément de fonction personnalisée dans Excel sur le Web ou Windows. Vous serez peut-être invité à ouvrir le volet Office du complément, bien que ce soit facultatif. Vous pouvez toujours exécuter vos fonctions personnalisées sans ouvrir le volet Office de votre complément.

# <a name="excel-on-windowstabexcel-windows"></a>[Excel sur Windows](#tab/excel-windows)

Pour tester votre complément dans Excel sous Windows, exécutez la commande suivante. Lorsque vous exécutez cette commande, le serveur Web local démarre et Excel s’ouvre avec votre complément chargé.

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[Excel sur le Web](#tab/excel-online)

Pour tester votre complément dans Excel sur le Web, exécutez la commande suivante. Lorsque vous exécutez cette commande, le serveur web local démarre.

```command&nbsp;line
npm run start:web
```

Pour utiliser votre complément de fonctions personnalisées, ouvrez un nouveau classeur dans Excel dans un navigateur. Dans ce classeur, effectuez les étapes suivantes pour chargement votre complément.

1. Dans Excel, sélectionnez l’onglet **insertion** , puis **compléments**.

   ![Insérer un ruban dans Excel sur le Web avec l’icône mes compléments mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)
   
2. Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.

3. Sélectionnez **Parcourir... ** et accédez au répertoire racine du projet créé par le Générateur de Yo Office.

4. Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.

---

## <a name="try-out-a-prebuilt-custom-function"></a>Essayer une fonction personnalisée prédéfinie

Le projet de fonctions personnalisées que vous avez créé à l’aide du générateur Yeoman contient certaines fonctions personnalisées prédéfinies, définies dans le fichier **./SRC/Functions/functions.js** . Le fichier **./manifest.xml** dans le répertoire racine du projet indique que toutes les fonctions personnalisées appartiennent à `CONTOSO` l’espace de noms.

Dans votre classeur Excel, essayez la `ADD` fonction personnalisée en procédant comme suit:

1. Sélectionnez une cellule et tapez `=CONTOSO`. Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans l’espace de noms `CONTOSO`.

2. Exécutez la `CONTOSO.ADD` fonction, en utilisant `10` des `200` nombres et comme paramètres d’entrée, en `=CONTOSO.ADD(10,200)` tapant la valeur dans la cellule et en appuyant sur entrée.

Le `ADD` fonction personnalisée calcule la somme des deux nombres que vous avez spécifiés comme paramètres d’entrée. La saisie de`=CONTOSO.ADD(10,200)` doit générer le résultat **210** dans la cellule une fois que vous appuyez sur ENTRÉE.

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé une fonction personnalisée dans un complément Excel! Ensuite, créez un complément plus complexe avec la fonctionnalité de diffusion de données en continu. Le lien suivant vous guide tout au long des étapes suivantes du didacticiel de complément Excel avec fonctions personnalisées.

> [!div class="nextstepaction"]
> [Didacticiel de complément de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a>Voir aussi

* [Vue d’ensemble des fonctions personnalisées](../excel/custom-functions-overview.md)
* [Métadonnées fonctions personnalisées](../excel/custom-functions-json.md)
* [Exécution de fonctions personnalisées Excel](../excel/custom-functions-runtime.md)
* [Meilleures pratiques de fonctions personnalisées](../excel/custom-functions-best-practices.md)
