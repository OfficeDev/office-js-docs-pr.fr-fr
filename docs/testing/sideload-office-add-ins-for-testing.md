---
title: Chargement de version test des compléments Office dans Office Online
description: Tester votre complément Office dans Office Online par chargement de version test
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 8870e955ca30c4a3b35f2b51e0e16a3ee634960d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451428"
---
# <a name="sideload-office-add-ins-in-office-online-for-testing"></a>Chargement de version test des compléments Office dans Office Online

Vous procéder à un chargement de version test pour installer un complément Office sans avoir à le placer au préalable dans un catalogue de compléments. Le chargement de version test s’effectue dans Office 365 ou Office Online. La procédure est légèrement différente entre les deux plateformes. 

Lorsque vous chargez une version test d’un complément, le manifeste du complément est stocké dans le stockage local du navigateur. Ainsi, si vous videz le cache du navigateur ou si vous basculez vers un autre navigateur, vous devez à nouveau charger une version test de complément.


> [!NOTE]
> Tel que décrit dans cet article, le chargement de version test est pris en charge dans Word, Excel et PowerPoint. Pour charger une version test de complément Outlook, voir la rubrique relative au [chargement de version test des compléments Outlook](/outlook/add-ins/sideload-outlook-add-ins-for-testing).

La vidéo suivante présente la procédure de chargement de version test de votre complément dans la version de bureau Office ou Office Online.  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-in-office-365"></a>Chargement de version test d’un complément Office dans Office 365


1. Connectez-vous à votre compte Office 365.
    
2. Ouvrez le lanceur d’applications à l’extrémité gauche de la barre d’outils et sélectionnez **Excel**,  **Word** ou **PowerPoint**, puis créez un document.
    
3. Ouvrez l’onglet **Insérer** dans le ruban, puis dans la section **Compléments**, choisissez **Compléments Office**.
    
4. Dans la boîte de dialogue **Compléments Office**, sélectionnez l’onglet **MON ORGANISATION**, puis **Télécharger mon complément**.
    
    ![Boîte de dialogue intitulée Complément Office avec un lien dans le coin supérieur gauche indiquant « Charger mon complément ».](../images/office-add-ins.png)

5.  **Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.
    
    ![Boîte de dialogue de chargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

6. Verify that your complément is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in the pane should appear.
    

## <a name="sideload-an-office-add-in-in-office-online"></a>Chargement de version test d’un complément Office dans Office Online


1. Ouvrez [Microsoft Office Online](https://office.live.com/).
    
2. Dans **Commencer à utiliser les applications en ligne maintenant**, choisissez **Excel**, **Word** ou **PowerPoint**, puis ouvrez un document.
    
3. Ouvrez l’onglet **Insérer** dans le ruban, puis dans la section **Compléments**, choisissez **Compléments Office**.
    
4. Dans la boîte de dialogue **Compléments Office**, sélectionnez l’onglet **MES COMPLÉMENTS**, choisissez **Gérer mes compléments**, puis **Télécharger mon complément**.
    
    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit indiquant « Gérer mes compléments » et une autre liste déroulante sous cette dernière avec l’option « Charger mon complément »](../images/office-add-ins-my-account.png)

5.  **Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.
    
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

6. Vérifiez que votre complément est installé. S’il s’agit d’une commande de complément, elle doit apparaître dans le ruban ou dans le menu contextuel. S’il s’agit d’un complément du volet Office, le volet doit apparaître.

> [!NOTE]
>Pour tester votre complément Office avec Edge, deux étapes de configuration sont nécessaires : 
>
> - Depuis une invite de commandes Windows, exécutez la ligne suivante : `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`
>
> - Entrez « **about:flags** » dans la barre de recherche Edge pour afficher les options des Paramètres de développeur.  Activez l’option « **Autoriser le bouclage localhost** », puis redémarrez Edge.

>    ![Option Autoriser le bouclage localhost de Edge avec la case à cocher activée.](../images/allow-localhost-loopback.png)

## <a name="sideload-an-add-in-when-using-visual-studio"></a>Chargement d’une version test d’un complément lors de l’utilisation de Visual Studio

Si vous développez votre complément à l’aide de Visual Studio, le processus de chargement d’une version de teste est similaire. La seule différence est que vous devez mettre à jour la valeur de l’élément **SourceURL** dans votre manifeste afin d’inclure l’URL complète de déploiement du complément.

> [!NOTE]
> Si vous pouvez charger une version test des compléments à partir de Visual Studio vers Office Online, vous ne pouvez pas les déboguer à partir de Visual Studio. Pour déboguer, vous devrez utiliser les outils de débogage du navigateur. Pour plus d’informations, voir [Débogage de compléments dans Office Online](debug-add-ins-in-office-online.md).

1. Dans Visual Studio, affichez la fenêtre **Propriétés** en choisissant **Affichage** -> **Fenêtre Propriétés**.
2. Dans l’**Explorateur de solutions**, sélectionnez le projet web. Cela a pour effet d’afficher les propriétés du projet dans la fenêtre **Propriétés**.
3. Dans la fenêtre Propriétés, copiez l’**URL SSL**.
4. Dans le projet de complément, ouvrez le fichier XML de manifeste. Veillez à modifier le code XML source. Pour certains types de projets, Visual Studio ouvre un affichage visuel du code XML qui ne fonctionnera pas pour l’étape suivante.
5. Cherchez toutes les instances de **~remoteAppUrl/** et remplacez-les par l’URL SSL que vous venez de copier. Vous verrez plusieurs remplacements en fonction du type de projet, et les nouvelles URL ressembleront à `https://localhost:44300/Home.html`.
6. Enregistrez le fichier XML.
7. Cliquez avec le bouton droit sur le projet web, puis sélectionnez **Déboguer** -> **Démarrer une nouvelle instance**. Cela a pour effet d’exécuter le projet web sans lancer Office.
8. À partir d’Office Online, chargez la version test du complément en suivant les étapes décrites précédemment dans [Chargement de version test d’un complément Office dans Office Online](#sideload-an-office-add-in-in-office-online).
