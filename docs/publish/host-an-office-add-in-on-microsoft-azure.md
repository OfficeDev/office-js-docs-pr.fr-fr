---
title: Héberger un complément pour Office sur Microsoft Azure | Microsoft Docs
description: Découvrez comment déployer une application web de complément sur Azure et charger une version test du complément pour le tester dans une application cliente Office.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: abe0012861a4c401f003704644fb9f530220521d
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292383"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a>Héberger un complément pour Office sur Microsoft Azure

Le complément Office le plus simple est composé d’un fichier manifeste XML et d’une page HTML. Le fichier manifeste XML décrit les caractéristiques du complément, telles que son nom, les clients de bureau Office qu’il peut exécuter, ainsi que l’URL de la page HTML du complément. La page HTML est contenue dans une application Web avec laquelle les utilisateurs interagissent lors de l’installation et l’exécution de votre complément dans une application cliente Office. Vous pouvez héberger l’application Web d’un complément Office sur une plateforme d’hébergement Web, y compris Azure.

Cet article décrit comment déployer une application web de complément sur Azure et [charger une version test du complément](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) pour le tester dans une application cliente Office.

## <a name="prerequisites"></a>Conditions préalables 

1. Installez [Visual Studio 2019](https://www.visualstudio.com/downloads) et choisissez d’inclure la charge de travail de **développement Azure**.

    > [!NOTE]
    > Si vous avez déjà installé Visual Studio 2019, [utilisez le programme d’installation Visual Studio Installer](/visualstudio/install/modify-visual-studio) pour vous assurer que la charge de travail de **développement Azure** est installée. 

2. Installation d’Office.

    > [!NOTE]
    > Si vous n’avez pas encore Office, vous pouvez vous [inscrire pour obtenir un essai gratuit d’un mois](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).

3. Obtenez un abonnement Azure.

    > [!NOTE]
    > Si vous n’avez pas encore d’abonnement Azure, vous pouvez [en obtenir un dans le cadre de votre abonnement Visual Studio](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/) ou vous [inscrire pour obtenir une version d’évaluation gratuite](https://azure.microsoft.com/pricing/free-trial). 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a>Étape 1 : Créer un dossier partagé pour héberger le fichier manifeste XML de votre complément

1. Ouvrez l’explorateur de fichiers sur votre ordinateur de développement.

2. Cliquez avec le bouton droit de la souris sur le lecteur C:\, puis choisissez **Nouveau** > **Dossier**.

3. Nommez le nouveau dossier AddinManifests.

4. Cliquez avec le bouton droit de la souris sur le dossier AddinManifests, puis choisissez **Partager avec** > **Des personnes spécifiques**.

5. Dans **Partage de fichiers**, sélectionnez la flèche déroulante vers le bas, puis choisissez **Tout le monde** > **Ajouter** > **Partager**.

> [!NOTE]
> Dans cette procédure, vous utilisez un partage de fichiers local en tant que catalogue approuvé où vous allez stocker le fichier manifeste XML du complément. Dans un scénario réel, vous pouvez choisir de [déployer le fichier manifeste XML dans un catalogue SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) ou de [publier le complément dans AppSource](/office/dev/store/submit-to-appsource-via-partner-center), à la place.

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a>Étape 2 : Ajouter le partage de fichiers au catalogue de compléments approuvés

1. Démarrez Word et créez un document.

    > [!NOTE]
    > Bien que cet exemple utilise Word, vous pouvez utiliser n’importe quelle application Office qui prend en charge des compléments Office comme Excel, Outlook, PowerPoint ou Project.

2. Choisissez **Fichier**  >  **Options**.

3. Dans la boîte de dialogue **Options Word**, choisissez **Centre de gestion de la confidentialité**, puis **Paramètres du Centre de gestion de la confidentialité**.

4. Dans la boîte de dialogue **Centre de gestion de la confidentialité**, choisissez **Catalogues de compléments approuvés**. Saisissez le chemin d’accès UNC (Universal Naming Convention) pour le partage de fichiers que vous avez créé précédemment en tant qu’**URL du catalogue** (par exemple, \\\NomDeVotreOrdinateur\AddinManifests), puis choisissez **Ajouter un catalogue**. 

5. Activez la case **Afficher dans le menu**.

    > [!NOTE]
    > Lorsque vous stockez un fichier manifeste XML de complément sur un partage qui est défini comme un catalogue de compléments web approuvés, le complément apparaît sous **Dossier partagé** dans la boîte de dialogue **Compléments Office** lorsque l’utilisateur accède à l’onglet **Insérer** dans le ruban et choisit **Mes compléments**.

6. Fermez Word.

## <a name="step-3-create-a-web-app-in-azure-using-the-azure-portal"></a>Étape 3 : Créer une application web dans Azure à l’aide du Portail Microsoft Azure

Pour créer l’application web à l’aide du portail Azure, procédez comme suit.

1. Connectez-vous au [portail Azure](https://portal.azure.com/) à l’aide de vos informations d’identification Azure.

2. Sous **Azure services**, sélectionnez **Applications web **.

3. Dans la page **Service d’applications**, sélectionnez **Ajouter**. Fournissez ces informations :

      - Choisissez l’**abonnement** à utiliser pour créer ce site.
      
      - Choisissez le **groupe de ressources** pour votre site. Si vous créez un groupe, vous devez également le nommer.
      
      - Entrez un **nom d’application** unique pour votre site. Azure vérifie que le nom du site est unique dans le domaine apps.net azureweb.

      - Indiquez si vous souhaitez publier à l'aide d'un code ou d'un conteneur docker.

      - Spécifiez une **pile d’exécution**.

      - Choisissez le **système d’exploitation** de votre site.

      - Choisissez une **Région**.

      - Choisissez le **plan de service d’applications** à utiliser pour créer ce site.

      - Sélectionnez **Créer**.

4. La page suivante vous indique que votre déploiement est en cours et quand il prend fin. Une fois l’opération terminée, sélectionnez **Accéder à la ressource**.  

5. Dans la section **Vue d’ensemble**, choisissez l’URL qui est affichée sous **URL**. Votre navigateur s’ouvre et affiche une page web avec le message « Votre application Service d’applications est opérationnelle. »

    > [!IMPORTANT]
    > Les sites web Azure [!include[HTTPS guidance](../includes/https-guidance.md)] fournissent automatiquement un point de terminaison HTTPS.

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a>Étape 4 : Créer un complément Office dans Visual Studio.

1. Démarrez Visual Studio en tant qu’administrateur.

2. Choisissez **Créer un nouveau projet**.

3. À l’aide de la zone de recherche, entrez **complément**.

4. Choisissez **Complément Word web** comme type de projet, puis cliquez sur **Suivant** pour accepter les paramètres par défaut.

Visual Studio crée un complément Word de base que vous pourrez publier tel quel, sans apporter de modifications à son projet web. Pour créer un complément pour une autre application Office, telle qu’Excel, répétez les étapes et choisissez un type de projet avec l’application Office souhaitée.

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a>Étape 5 : Publier votre application web de complément Office sur Azure

1. Avec votre projet de complément ouvert dans Visual Studio, développez le nœud de solutions dans **Explorateur de solutions**, puis sélectionnez **Service d’applications**.

2. Cliquez avec le bouton droit de la souris sur le projet web, puis choisissez **Publier**. Le projet web contient les fichiers d’application web du complément Office, et il s’agit donc du projet que vous publiez sur Azure.

3. Sur l’onglet **Publier** :

      - Choisissez **Microsoft Azure Application Service**.

      - Choisissez **Sélectionner**.

      - Choisissez **Publier**.

4. Visual Studio publie le projet web pour votre complément Office sur votre site web Azure. Une fois le projet web publié par Visual Studio, votre navigateur s’ouvre et affiche une page web avec le texte « Votre application de service d’application a été créée. » Il s’agit de la page active par défaut pour l’application web.

5. Copiez l’URL racine (par exemple : https://YourDomain.azurewebsites.net) ; vous en aurez besoin lorsque vous modifierez le fichier manifeste de complément plus loin dans cet article.

## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a>Étape 6 : Modifier et déployer le fichier manifeste XML

1. Dans Visual Studio avec l’exemple de complément Office ouvert dans l’**explorateur de solutions**, développez la solution pour que les deux projets s’affichent.

2. Développez le projet macro complémentaire Office (par exemple WordWebAddIn), le dossier manifeste d’avec le bouton droit de la souris et sélectionnez **Ouvrir**. Le fichier manifeste XML du complément s’ouvre.

3. Dans le fichier manifeste XML, recherchez et remplacez toutes les instances de « ~remoteAppUrl » par l’URL racine de l’application web du complément sur Azure. Il s’agit de l’URL que vous avez copiée précédemment une fois que vous avez publié l’application web du complément sur Azure (par exemple : https://YourDomain.azurewebsites.net). 

4. Choisissez **Fichier**, puis **Enregistrer tout**. Ensuite, copiez le fichier manifeste XML du complément (par exemple, WordWebAddIn.xml).

5. À l’aide du programme **Explorateur de fichier**, accédez au partage de fichiers réseau que vous avez créé à l’[Étape 1 : Créer un dossier partagé](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file), puis collez le fichier manifeste dans le dossier.

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a>Étape 7 : insérer et exécuter le complément dans l’application cliente Office

1. Démarrez Word et créez un document.

2. Sur le ruban, cliquez sur **Insérer** > **Mes compléments**.

3. Dans la boîte de dialogue **Compléments Office**, choisissez **DOSSIER PARTAGÉ**. Word recherche le dossier que vous avez désigné comme catalogue de compléments approuvés (à l’[étape 2 : Ajouter le partage de fichiers au catalogue de compléments approuvés](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) et affiche les compléments dans la boîte de dialogue. Vous devriez voir l’icône de votre exemple de complément.

4. Cliquez sur l’icône de votre complément, puis choisissez **Ajouter**. Un bouton **Afficher le volet de tâches** pour votre complément est ajouté au ruban.

5. Dans le ruban de l’onglet **Accueil**, choisissez le bouton **Afficher le volet de tâches**. Le complément s’ouvre dans un volet de tâches à droite du document actif.

6. Vérifiez que le complément fonctionne en sélectionnant du texte dans le document et en choisissant le bouton **Mettre en surbrillance** dans le volet de tâches.

## <a name="see-also"></a>Voir aussi

- [Publier votre complément Office](../publish/publish.md)
- [Publier votre complément à l’aide de Visual Studio](../publish/package-your-add-in-using-visual-studio.md)
