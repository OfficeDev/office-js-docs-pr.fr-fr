---
title: "Chargement de version test des compléments\_Office dans Office\_Online"
description: ''
ms.date: 12/04/2017
---

# <a name="sideload-office-add-ins-in-office-online-for-testing"></a>Chargement de version test des compléments Office dans Office Online

Vous pouvez installer un complément Office test sans avoir à le placer au préalable dans un catalogue de compléments en utilisant le chargement de version test. Le chargement de version test peut être effectué sur Office 365 ou Office Online. La procédure présente de légères différences d’une plateforme à l’autre. 

Lorsque vous chargez une version test d’un complément, le manifeste du complément est stocké dans le stockage local du navigateur. Ainsi, si vous videz le cache du navigateur ou si vous basculez vers un autre navigateur, vous devez à nouveau charger une version test de complément.


> [!NOTE]
> Tel que décrit dans cet article, le chargement de version test est pris en charge dans Word, Excel et PowerPoint. Pour charger une version test de complément Outlook, voir la rubrique relative au [chargement de version test des compléments Outlook](https://docs.microsoft.com/fr-fr/outlook/add-ins/sideload-outlook-add-ins-for-testing).

La vidéo suivante présente la procédure de chargement de version test de votre complément dans la version de bureau Office ou Office Online.  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-on-office-365"></a>Chargement de version test d’un complément Office dans Office 365


1. Connectez-vous à votre compte Office 365.
    
2. Ouvrez le lanceur d’applications à l’extrémité gauche de la barre d’outils et sélectionnez **Excel**,  **Word** ou **PowerPoint**, puis créez un document.
    
3. Ouvrez l’onglet **Insérer** dans le ruban, puis dans la section **Compléments**, choisissez **Compléments Office**.
    
4. Dans la boîte de dialogue **Compléments Office**, sélectionnez l’onglet **MON ORGANISATION**, puis **Télécharger mon complément**.
    
    ![Boîte de dialogue intitulée Complément Office avec un lien dans le coin supérieur gauche indiquant « Charger mon complément ».](../images/office-add-ins.png)

5.  **Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.
    
    ![Boîte de dialogue de chargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

6. Verify that your complément is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in the pane should appear.
    

## <a name="sideload-an-office-add-in-on-office-online"></a>Charger une version test d’un complément Office sur Office Online


1. Open [Microsoft Office Online](https://office.live.com/).
    
2. Dans **Commencer à utiliser les applications en ligne maintenant**, choisissez **Excel**, **Word** ou **PowerPoint**, puis ouvrez un document.
    
3. Ouvrez l’onglet **Insérer** dans le ruban, puis dans la section **Compléments**, choisissez **Compléments Office**.
    
4. Dans la boîte de dialogue **Compléments Office**, sélectionnez l’onglet **MES COMPLÉMENTS**, choisissez **Gérer mes compléments**, puis **Télécharger mon complément**.
    
    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit indiquant « Gérer mes compléments » et une autre liste déroulante sous cette dernière avec l’option « Charger mon complément »](../images/office-add-ins-my-account.png)

5.  **Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.
    
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

6. Vérifiez que votre complément est installé. S’il s’agit d’une commande de complément, elle doit apparaître dans le ruban ou dans le menu contextuel. S’il s’agit d’un complément du volet Office, le volet doit apparaître.

## <a name="sideload-an-add-in-when-using-visual-studio"></a>Chargement d’une version test d’un complément lors de l’utilisation de Visual Studio

Si vous développez votre complément à l’aide de Visual Studio, le processus de chargement d’une version de teste est similaire. La seule différence est que vous devez mettre à jour la valeur de l’élément **SourceURL** dans votre manifeste, de sorte à inclure l’URL entière de l’emplacement de déploiement du complément. 

Si vous êtes en train de développer votre complément, recherchez-le dans le fichier manifest.xml et mettez à jour la valeur de l’élément **SourceLocation** de façon à inclure un URI absolu. Visual Studio met en place un jeton pour votre déploiement localhost.

Par exemple : 

```xml
<SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
```
