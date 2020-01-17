---
title: Utiliser le générateur Yeoman pour créer un complément Office qui utilise l’authentification unique (aperçu)
description: Utiliser le générateur Yeoman pour créer un complément Office Node.js qui utilise l’authentification unique (aperçu)
ms.date: 01/13/2020
ms.prod: non-product-specific
localization_priority: Priority
ms.openlocfilehash: 1f02f03fec0d6be32fc7a0d6b98fce30e19c28e2
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217364"
---
# <a name="use-the-yeoman-generator-to-create-an-office-add-in-that-uses-single-sign-on-preview"></a>Utiliser le générateur Yeoman pour créer un complément Office qui utilise l’authentification unique (aperçu)

Dans cet article, vous allez découvrir le processus d’utilisation du générateur Yeoman pour créer un complément Office pour Excel, Word ou PowerPoint qui utilise l’authentification unique (SSO) lorsque c’est possible, et utilise une autre méthode d’authentification utilisateur lorsque l’authentification unique n’est pas prise en charge.

> [!TIP]
> Avant d'essayer de terminer ce démarrage rapide, consultez la section [Activer l'authentification unique pour les compléments Office](../develop/sso-in-office-add-ins.md) pour apprendre les concepts de base de l'authentification unique dans les compléments Office. 
 
Le générateur Yeoman simplifie le processus de création d’un complément d’authentification unique en automatisant les étapes nécessaires pour configurer l’authentification unique dans Azure et la génération du code nécessaire pour qu’un complément utilise l’authentification unique. Si vous souhaitez avoir une description détaillée de la procédure à suivre pour effectuer manuellement les étapes que le générateur Yeoman automatise, veuillez consulter le didacticiel [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md).

## <a name="prerequisites"></a>Conditions préalables

* [Node.js](https://nodejs.org) (version 10.15.0 ou ultérieure)

* La dernière version de[Yeoman](https://github.com/yeoman/yo) et de [Yeoman Générateur de compléments Office](https://github.com/OfficeDev/generator-office). Pour installer ces outils globalement, exécutez la commande suivante à partir de l’invite de commande :

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="create-the-add-in-project"></a>Création du projet de complément

> [!TIP]
> Le générateur Yeoman peut créer un complément Office avec authentification unique pour Excel, Word ou PowerPoint, et peut être créé avec des scripts de type JavaScript ou TypeScript. Les instructions suivantes indiquent `JavaScript` et `Excel`, mais vous devez choisir le type de script et l’application client Office les mieux adaptées à votre scénario.

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Sélectionnez un type de projet :** `Office Add-in Task Pane project supporting single sign-on`
- **Sélectionnez un type de script :** `Javascript`
- **Comment souhaitez-vous nommer votre complément ?** `My SSO Office Add-in`
- **Quelle application client Office voulez-vous prendre en charge ?** `Excel`

![Capture d’écran des invites et des réponses relatives au générateur Yeoman](../images/yo-office-sso-excel.png)

Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Explorer le projet

Le projet de complément que vous avez créé à l’aide du générateur Yeoman contient un code pour un complément de volet Office avec authentification unique.

- Le fichier **./manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.

- Le fichier **./src/taskpane/taskpane.html** contient les balises HTML du volet Office.
- Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.
- Le fichier **./src/taskpane/taskpane.js** contient le code de l’API JavaScript pour Office qui facilite l’interaction entre le volet Office et l’application hôte Office.

- Le fichier **./src/helpers/documentHelper.js** utilise la bibliothèque JavaScript Office pour ajouter les données de Microsoft Graph au document Office.
- Le fichier **./src/helpers/fallbackauthdialog.html** est la page sans interface utilisateur qui charge le code JavaScript de la méthode d’authentification de secours.
- Le fichier **./src/helpers/fallbackauthdialog.html** contient le code JavaScript de la méthode d’authentification de secours qui se connecte l'utilisateur avec msal.js.
- Le fichier **./SRC/helpers/fallbackauthhelper.js** contient le volet Office JavaScript qui appelle la méthode d’authentification de secours dans les scénarios lorsque l’authentification unique n’est pas prise en charge.
- Le fichier **./src/helpers/ssoauthhelper.js** contient l’appel JavaScript à l’API de l’authentification unique, `getAccessToken`, reçoit le jeton d’amorçage, initialise le remplacement du jeton d’amorçage pour un jeton d’accès à Microsoft Graph et appelle Microsoft Graph pour les données.

- Le fichier **./ENV** dans le répertoire racine du projet définit les constantes utilisées par le projet de complément.
    > [!NOTE]
    > Certaines des constantes définies dans ce fichier sont utilisées pour simplifier le processus d’authentification unique. Vous pouvez mettre à jour les valeurs de ce fichier pour qu'elles correspondent à votre scénario spécifique. Par exemple, vous pouvez mettre à jour ce fichier pour spécifier une autre étendue, si votre complément nécessite une autre valeur que `User.Read`.

## <a name="configure-sso"></a>Configurer l’authentification unique

À ce stade, votre projet de complément a été créé et contient le code nécessaire pour simplifier le processus d’authentification unique. Ensuite, procédez comme suit pour configurer l’authentification unique pour votre complément.

1. Accédez au dossier racine du projet.

    ```command&nbsp;line
    cd "My SSO Office Add-in"
    ```

2. Exécutez la commande suivante pour configurer l’authentification unique pour le complément.

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > Cette commande échouera si votre locataire est configuré pour nécessiter une authentification à deux facteurs. Dans ce scénario, vous devez effectuer manuellement les étapes d’inscription et de configuration de l’authentification unique de l’application Azure, comme décrit dans le didacticiel [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md).

3. Une fenêtre de navigateur web s’ouvre et vous invite à vous connecter à Azure. Connectez-vous à Azura à l’aide de vos informations d’identification d’administrateur Office 365. Ces informations d’identification sont utilisées pour inscrire une nouvelle application dans Azure et configurer les paramètres requis par l’authentification unique.

    > [!NOTE]
    > Si vous vous connectez à Azure à l’aide d’informations d’identification non-administrateur au cours de cette étape, le script `configure-sso` ne peut pas fournir d’accord d’administrateur pour le complément aux utilisateurs au sein de votre organisation. Par conséquent, l’authentification unique ne sera pas disponible pour les utilisateurs du complément. vous serez invité à vous connecter.

4. Une fois vos informations d'identification saisies, fermez la fenêtre du navigateur et revenez à l'invite de commande. Au fur et à mesure du processus de configuration de l’authentification unique, les messages d’État s’affichent sur la console. Comme décrit dans la section messages de la console, les fichiers du projet de complément que le générateur Yeoman a créé sont automatiquement mis à jour avec les données requises par le processus d’authentification unique.

## <a name="try-it-out"></a>Essayez

1. Une fois le processus de configuration de l’authentification unique terminé, exécutez la commande suivante pour créer le projet, démarrez le serveur web local et mettez votre complément en sideload dans l’application client Office précédemment sélectionnée.

    > [!NOTE]
    > Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez. Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman.

    ```command&nbsp;line
    npm start
    ```

2. Dans l’application client Office qui s’ouvre lorsque vous exécutez la commande précédente (par exemple, Excel, Word ou PowerPoint), assurez-vous que vous êtes connecté avec un utilisateur membre de la même organisation Office 365 que le compte d’administrateur Office 365 que vous avez utilisé pour vous connecter à Azure lors de la configuration de l’authentification unique à l’étape 3 de la [section précédente](#configure-sso). Cette opération permet d’établir les conditions appropriées pour la réussite de l’authentification unique. 

3. Dans l’application client Office, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément. L’image ci-après illustre ce bouton dans Excel.

    ![Bouton Complément Excel](../images/excel-quickstart-addin-3b.png)

4. Au bas du volet Office, sélectionnez le bouton **Obtenir mes informations de profil utilisateur** pour lancer le processus d’authentification unique. 

    > [!NOTE] 
    > Si vous n’êtes pas encore connecté à Office à ce stade, vous serez invité à le faire. Comme décrit précédemment, vous devez vous connecter avec un utilisateur membre de la même organisation Office 365 que le compte d’administrateur Office 365 que vous avez utilisé pour vous connecter à Azure lors de la configuration de l’authentification unique à l’étape 3 de la [section précédente](#configure-sso), si vous voulez que l’authentification unique soit réussie.

5. Si une boîte de dialogue s’affiche pour demander des autorisations pour le compte du complément, cela signifie que l’authentification unique n’est pas prise en charge pour votre scénario et que le complément est plutôt repassé à une autre méthode d’authentification des utilisateurs. Cela peut se produire lorsque l’administrateur du locataire n’a pas accordé le consentement du complément pour accéder à Microsoft Graph, ou lorsque l’utilisateur n’est pas connecté à Office à l’aide d’un compte Microsoft valide ou d’un compte Office 365 (« professionnel ou scolaire »). Sélectionnez le bouton **Accepter** dans la fenêtre de boîte de dialogue pour continuer.

    ![Boîte de dialogue demande d’autorisation](../images/sso-permissions-request.png)

    > [!NOTE]
    > Une fois qu’un utilisateur a accepté cette demande d’autorisation, il n’est plus invité à le faire à l’avenir.

6. Le complément récupère les informations de profil de l’utilisateur connecté et écrit celui-ci dans le document. L’image suivante montre un exemple d’informations de profil écrites dans une feuille de calcul Excel.

    ![Informations de profil utilisateur dans la feuille de calcul Excel](../images/sso-user-profile-info-excel.png)

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé un complément de volet Office qui utilise l’authentification unique lorsque c’est possible, et utilise une autre méthode d’authentification utilisateur lorsque l’authentification unique n’est pas prise en charge. Pour en savoir plus sur les étapes de configuration de l’authentification unique effectuées automatiquement par le générateur Yeoman et le code facilitant le processus d’authentification unique, veuillez consultez le didacticiel [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md).

## <a name="see-also"></a>Consultez aussi

- [Activer l’authentification unique pour des compléments Office](../develop/sso-in-office-add-ins.md)
- [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md)
- [Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)](../develop/troubleshoot-sso-in-office-add-ins.md)