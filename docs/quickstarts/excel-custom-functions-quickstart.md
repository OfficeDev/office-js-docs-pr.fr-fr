---
ms.date: 03/06/2019
description: Développement de fonctions personnalisées dans le Guide de démarrage rapide d’Excel.
title: Démarrage rapide des fonctions personnalisées (aperçu)
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 3ea7ec4c2089aaa4e9f193a45e7c4a31c691f213
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/01/2019
ms.locfileid: "33517071"
---
# <a name="get-started-developing-excel-custom-functions"></a>Prise en main du développement de fonctions personnalisées Excel

Avec les fonctions personnalisées, les développeurs peuvent désormais ajouter de nouvelles fonctions à Excel en les définissant en JavaScript ou en une machine à écrire dans le cadre d’un complément. Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe `SUM()`quelle fonction native dans Excel, comme.

## <a name="prerequisites"></a>Conditions préalables

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Vous aurez besoin des outils et ressources connexes suivants pour commencer à créer des fonctions personnalisées.

- [Node.js](https://nodejs.org/en/) (version 8.0.0 ou ultérieure)

- [Git Bash](https://git-scm.com/downloads) (ou un autre client Git)

- La dernière version de[Yeoman](https://yeoman.io/) et de [Yeoman Générateur de compléments Office](https://www.npmjs.com/package/generator-office). Pour installer ces outils globalement, exécutez la commande suivante à partir de l’invite de commande :

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > Même si vous avez déjà installé le générateur Yeoman, nous vous recommandons de mettre à jour votre package vers la dernière version à partir de NPM.

## <a name="build-your-first-custom-functions-project"></a>Création de votre premier projet de fonctions personnalisées

Pour commencer, vous utiliserez le Yeoman Générateur pour créer le projet de fonctions personnalisées. Cette option définit votre projet, avec la structure de dossiers correct, les fichiers source et les dépendances pour commencer le codage de vos fonctions personnalisées.

1. Exécutez la commande suivante, puis répondez aux invitations comme suit.

    ```command&nbsp;line
    yo office
    ```

    - Choisissez un type de projet : `Excel Custom Functions Add-in project (...)`

    - Choisissez un type de script : `JavaScript`

    - Comment souhaitez-vous nommer votre complément ? `stock-ticker`

    ![Le générateur de yeoman pour les compléments Office vous invite pour les fonctions personnalisées](../images/12-10-fork-cf-pic.jpg)

    Le générateur crée le projet et installe les composants Node.js de la prise en charge.

2. Naviguez jusqu’au dossier de projet que vous venez de créer.

    ```command&nbsp;line
    cd stock-ticker
    ```

3. Approuvez le certificat auto-signé dont vous avez besoin pour exécuter ce projet. Pour obtenir des instructions détaillées pour Windows ou Mac, voir [Ajout des Certificats Auto-signés comme Certificat Racine Approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).  

4. Construire le projet.

    ```command&nbsp;line
    npm run build
    ```

5. Démarrez le serveur web local qui est exécuté dans Node.js.

    - Si vous utilisez Excel pour Windows pour tester vos fonctions personnalisées, exécutez la commande suivante pour démarrer le serveur Web local, lancez Excel et chargement le complément:

        ```command&nbsp;line
         npm run start
        ```
        Après avoir exécuté cette commande, votre invite de commandes affiche des détails sur le démarrage du serveur Web. Excel commence avec votre complément chargé. Si vous complément ne charge pas, vérifiez que vous avez correctement terminé l’étape 3.

    - Si vous utilisez Excel Online pour tester vos fonctions personnalisées, exécutez la commande suivante pour démarrer le serveur Web local:

        ```command&nbsp;line
        npm run start-web
        ```

         Après avoir exécuté cette commande, votre invite de commandes affiche des détails sur le démarrage du serveur Web. Pour utiliser vos fonctions, ouvrez un nouveau classeur dans Excel online. Dans ce classeur, vous devrez charger votre complément. 

        Pour ce faire, sélectionnez l’onglet **Insérer** sur le ruban et sélectionnez **Get Add-ins**. Dans la nouvelle fenêtre qui s’affiche, vérifiez que vous êtes dans l’onglet **mes compléments** . Ensuite, sélectionnez **gérer mes compléments _GT_ Télécharger mon complément**. Recherchez votre fichier manifeste et téléchargez-le. Si votre complément ne se charge pas, vérifiez que vous avez correctement terminé l’étape 3.

## <a name="try-out-the-prebuilt-custom-functions"></a>Tester les fonctions personnalisées prédéfinies

Le projet de fonctions personnalisées que vous avez créé à l’aide du générateur Office Yo contient certaines fonctions personnalisées prédéfinies, définies dans le fichier **src/customfunction.js**. Le fichier**manifest.xml**dans le répertoire racine du projet indique que toutes les fonctions personnalisées appartiennent à l’ `CONTOSO` espace de noms.

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
